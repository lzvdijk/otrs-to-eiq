"""
Mailbox poller for OTRS/EIQ integration

This script queries a mailbox to check if there are new e-mails
If an e-mail is found, the matching ticket is retrieved from OTRS
The ticket can then be used to create an 'Incident' entity in EIQ

"""

from email import parser

import email
import imaplib
import poplib
import json
import re
import configparser
import requests
import structlog

from urllib.parse import urlsplit

from eiqlib.eiqjson import EIQEntity, EIQRelation
from eiqlib.eiqcalls import EIQApi

# immediately initialise an error log
log = structlog.get_logger(__name__)
config = configparser.ConfigParser()

try:
    config.read("config.ini")
except Exception:
    structlog.get_logger().error("Config file not found, please create one \
                                    under ./config.ini")

def init_api(eiq_url, eiq_user, eiq_pass):
    """
    Setup a connection to EIQ using eiqlib
    """
    api_inst = EIQApi()
    api_inst.set_host(eiq_url)
    api_inst.set_credentials(eiq_user, eiq_pass)

    return api_inst

def get_ticket_date(ticket):
    """
    Return the last modification date of a ticket as the report date
    """
    ctdate = ""
    chdate = ""
    crdate = ""
    cldate = ""

    if 'Article' in ticket['Ticket'][0]:
        if 'ChangeTime' in ticket['Ticket'][0]['Article'][-1]:
            ctdate = ticket['Ticket'][0]['Article'][-1]['ChangeTime']
    if 'Created' in ticket['Ticket'][0]:
        crdate = ticket['Ticket'][0]['Created']
    if 'Changed' in ticket['Ticket'][0]:
        chdate = ticket['Ticket'][0]['Changed']
    if 'Closed' in ticket['Ticket'][0]:
        cldate = ticket['Ticket'][0]['Closed']

    return max(ctdate, chdate, crdate, cldate)

def get_ticket_content(ticket):
    """
    Concat all relevant information from an OTRS ticket as the report description.
    Checks for placeholder strings.
    """
    descriptionstring = ""

    if 'Owner' in ticket['Ticket'][0]:
        descriptionstring = descriptionstring + "Current ticket owner: " \
                            + ticket['Ticket'][0]['Owner'] + "<br />"

    descriptionstring = descriptionstring + "Link to the original ticket here: " \
                        + "https://" + config['OTRS']['otrs_server'] \
                        + "/otrs/index.pl?Action=AgentTicketZoom;TicketID=" \
                        + ticket['Ticket'][0]['TicketID'] + "<br />"

    for field in ticket['Ticket'][0]['DynamicField']:
        if field['Name'] == 'LaaSta' and field['Value'] != None:
            if "Verander de Title hierboven naar iets wat het incident duidelijk weergeeft" \
            not in field['Value']:
                descriptionstring = descriptionstring \
                                    + "Laatste status: <br />" \
                                    + field['Value'] + "<br />"
        if field['Name'] == 'onderzoek' and field['Value'] != None:
            if "This report has also been sent to the KPN CISO CERT who will track your progress" \
            not in field['Value']:
                descriptionstring = descriptionstring + "Onderzoek: " + field['Value'] + "<br />"
        if field['Name'] == 'TicketSystems' and field['Value'] != None:
            descriptionstring = descriptionstring \
                                + "Betrokken systemen: " \
                                + field['Value'] + "<br />"
    # any other fields we want to display in the report?

    # test for free text fields in ticket articles
    if 'Article' in ticket['Ticket'][0]:
        if 'Body' in ticket['Ticket'][0]['Article'][-1]:
            descriptionstring = descriptionstring + "Laatste artikel: " \
                                + ticket['Ticket'][0]['Article'][-1]['Body'] + "<br />"

    # any other fields we want to display in the report?

    return descriptionstring

def get_ticket(ticketid):
    """
    Get a ticket from OTRS based on the OTRS ticket id (e.g. 34568)
    Note the difference between ticket id and ticket number.
    """
    # check if we have a ticketID
    if ticketid:
        if config['DEBUG']['verbose']:
            structlog.get_logger().info("Retrieving ticket at: "+config['OTRS']['otrs_address'] \
                  +"GetTicket/{}".format(ticketid))
        try:
            req = requests.post(config['OTRS']['otrs_address']+"GetTicket/{}".format(ticketid),
                                json={"UserLogin":config['OTRS']['otrs_user'],
                                      "Password":config['OTRS']['otrs_pass'],
                                      # retrieve _all_ the information
                                      "DynamicFields":"1",
                                      "Extended":"1",
                                      "AllArticles":"1",},
                                # DangerZone(tm)
                                verify=False)
        except Exception as err:
            structlog.get_logger().info("Error during EIQ call: "+err)

        # if we get a valid response, jsonify & return it.
        if not req.raise_for_status():
            return json.loads(req.content.decode('utf-8'))
        # otherwise return the error code
        return req.status_code()
    print("Invalid ticketID given!")

def search_ticket_number(searchvalue):
    """
    Search for a ticket in OTRS using a specific ticket NUMBER
    """
    # check if we actually have something to search for
    if searchvalue:
        if len(searchvalue) > 5:
            try:
                req = requests.post(config['OTRS']['otrs_address']+"SearchTicket/",
                                    json={"UserLogin":config['OTRS']['otrs_user'],
                                          "Password":config['OTRS']['otrs_pass'],
                                          "DynamicFields":"1",
                                          "TicketNumber":searchvalue},
                                    # SSL verification
                                    verify=False)
            except Exception as err:
                structlog.get_logger().info("Error during EIQ call: "+err)

            # if we get something cool, jsonify & return it
            if not req.raise_for_status():
                return json.loads(req.content.decode('utf-8'))
            # otherwise return the error code
            return req.status_code()
    structlog.get_logger().info("Invalid search query")

# def process_message_part(part):
#     if (part.get_content_maintype() == 'multipart'
#             or part.get('Content-Disposition') is None):
#         return

#     filename = part.get_filename()
#     if filename is None:
#         return None

#     content = part.get_payload(decode=True)
#     return content

def get_mail(mailserver, username, password):
    """
    Attempts to connect to an imap mailserver and return the titles of recent email messages

    These titles can be used to retrieve matching OTRS tickets (title search)
    """
    titles = []

    try:
        mail_con = imaplib.IMAP4_SSL(mailserver)
        mail_con.login(username, password)
        mail_con.select('INBOX')
        status, response = mail_con.search("UTF-8", "*")

        message_ids = response[0].split()

        for message_id in message_ids:

            status, data = mail_con.fetch(message_id, '(RFC822)')
            if status != 'OK':
                structlog.get_logger().error("Error fetching a message: {}".format(status))

            message = email.message_from_bytes(data[0][1])

            title = message['Subject']
            titles.append(title)

        return titles

    except Exception as err:
        structlog.get_logger().error("Error connecting to mailserver: "+err)
        return None

def check_exists(ticket, eiq_api):
    """
    Checks if an incident item exists in EIQ for a given OTRS ticket number
    """

    token = eiq_api.do_auth()

    # create necessary authorisation headers
    headers = {}
    headers['User-Agent'] = 'eiqlib/1.0'
    headers['Authorization'] = 'Bearer %s' % (token['token'],)
    headers['Cookie'] = 'platform-api-token=%s' % (token['token'],)

    # create a search query
    query = "\"CERT#"+ticket['Ticket'][0]['TicketNumber']+"\""

    if config['DEBUG']['verbose']:
        print("\nQuerying EIQ for: "+query)

    req = None

    try:
        req = eiq_api.do_call("/search-entities/", "POST", headers,
                              bytes(json.dumps({"query":{"bool":{"must":{"query_string":{\
                              "query":query, "time_zone":"+01:00", "lenient":"false"}}}},\
                              "aggregations":{"query_total":{"value_count":{"field":"_type"}},\
                              "type_counts":{"terms":{"field":"data.type", "order":{"_term":"asc"},\
                              "size":100}}, "source_counts":{"terms":{"field":"sources.name",\
                              "order":{"_term":"asc"}, "size":100}}, "tlp_counts":{"terms":{\
                              "field":"meta.tlp_color", "order":{"_term":"asc"}, "size":100,\
                              "missing":"NONE"}}, "source_reliability_counts":{"terms":{\
                              "field":"meta.source_reliability", "order":{"_term":"asc"},\
                              "size":100}}, "dataset_counts":{"terms":{"field":"intel_sets",\

                              "order":{"_term":"asc"}, "size":100}}}}), encoding="utf-8"))
    except Exception as err:
        structlog.get_logger().info("Error during EIQ call: "+str(err))

    # if we get a valid response, parse it.
    if req:
        if config['DEBUG']['verbose']:
                    print("EIQ returned response: "+str(req))
        if 'hits' in req:
            if req['hits']['total'] > 0:
                return True
            return False

    # otherwise return None
    return None

def create_incident(ticket, eiq_api):
    """
    Creates an incident entity in EIQ with information from a given OTRS ticket
    """

    incident = EIQEntity()

    incident.set_entity(incident.ENTITY_INCIDENT)
    incident.set_entity_title(ticket['Ticket'][0]['Title']+" - CERT#"+ticket['Ticket'][0]['TicketNumber']+" - [OTRS]")

    ticket_description = "[OTRS][Incident] CERT#"+ticket['Ticket'][0]['TicketNumber']

    ticket_description = ticket_description + "Ticket date: " \
                         + get_ticket_date(ticket) + " </br>"

    ticket_description = ticket_description + get_ticket_content(ticket)

    incident.set_entity_description(ticket_description)
    incident.set_entity_confidence(incident.CONFIDENCE_HIGH)
    incident.set_entity_tlp('AMBER')
    incident.add_ttp_type(incident.TTP_ADVANTAGE)
    incident.add_discovery_type(incident.DISCOVERY_UNKNOWN)
    incident.add_category_type(incident.CATEGORY_TEST)
    incident.set_entity_source(config['EIQ']['eiq_uid'])

    # sourceurl = "{0.scheme}://{0.netloc}".format(urlsplit(config['OTRS']['otrs_address'])) \
    #             +"/otrs/index.pl?Action=AgentTicketZoom;TicketID="+ticket['Ticket'][0]['TicketID']
    # incident.set_entity_source_reference(sourceurl)
    # incident.set_entity_source_description("OTRS Mail Feed")

    if config['DEBUG']['verbose']:
        print("\nGenerated entity:")
        print(incident.get_as_json())

    return incident.get_as_json()

def process_ticket(ticket, eiq_api):
    """
    Check if a given ticket exists in EIQ, and create an incident entity if it doesn't

    Future feature: merge new articles in existing incident entities based on ticket timestamps
    """

    if 'Error' in ticket:
        structlog.get_logger().info("Ticket error: "+ str(ticket['Error']))
        return None

    # check if an incident entity already exists in EIQ
    if not check_exists(ticket, eiq_api):
        incident = create_incident(ticket, eiq_api)
        if config['DEBUG']['verbose']: 
            print("\nCreating entity in EIQ\n")
        eiq_api.create_entity(incident)

def main():
    """
    Reads config file, gets latest e-mails from a mailbox, does OTRS stuff with results.
    """

    emails = get_mail(config['MAIL']['mail_address'], config['MAIL']['mail_user'],
                      config['MAIL']['mail_pass'])

    eiq_api = init_api(config['EIQ']['eiq_address'],
                       config['EIQ']['eiq_user'],
                       config['EIQ']['eiq_pass'])

    for title in emails:
        ticketnumber = re.search(r'CERT#(\d*)', title).group(1)
        if ticketnumber is not None:
            if config['DEBUG']['verbose']:
                structlog.get_logger().info("Found ticket #: "+ticketnumber)
            ticketid = search_ticket_number(ticketnumber)
            if 'TicketID' in ticketid:
                ticket = get_ticket(ticketid['TicketID'][0])
                if config['DEBUG']['verbose']:
                    structlog.get_logger().info("\nProcessing ticket #: "+ticket['Ticket'][0]['TicketNumber'])
                process_ticket(ticket, eiq_api)


if __name__ == "__main__":
    main()
