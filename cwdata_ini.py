#! python3
# -*- coding: utf-8 -*-

"""
Connectwise API Data Grabber

Need to enter API information in the apikey.ini file in the same directory as
the main.py file.
"""
import base64
import requests  # need to be installed
import datetime
import configparser

__author__ = 'Dustin Rowland'
__email__ = 'lee.rowland@gmail.com'
__copyright__ = '2015, Dustin Rowland'
__status__ = 'Development'

parsed_json = []
status_fu = {}
status_qa = {}
status_re = {}
status_wr = {}
status_et = {}
status_dis = {}
status_sch = {}
status_pro = {}
status_new = {}

# this is to read from the apikey.ini file using the configparser module
api_file = configparser.ConfigParser()
api_file.read('apikey.ini')
# first field in get is section title, second field is key value
companyName = api_file.get('CW API Key', 'companyName')
pubKey = api_file.get('CW API Key', 'publicKey')
privKey = api_file.get('CW API Key', 'privateKey')
# combine the key values to be used by b64_clientid function
client_id = companyName + '+' + pubKey + ':' + privKey


def b64_clientid(client_id):
    """
    Take the client_id from the apikey.ini file and
    convert to base64 for CW API authentication.
    """
    # encode the string in utf-8 (unicode) from ASCII
    cid_encode = client_id.encode('utf-8')
    # convert the encoded string in base64 bytes
    cid_base64 = base64.b64encode(cid_encode)
    # decode the base64 as utf-8 to remove the b' and \n
    cid_decode = cid_base64.decode('utf-8')

    return cid_decode


def cwJSON(section, subsection, page):
    """
    Pull JSON data from Connectwise using the REST API. This will pull the
    client_id from the b64_clientid function and pass it to the authentication
    header below.
    """
    # base string of the API URL, we can address all sections with this string
    API_base = ('https://api-na.myconnectwise.net/v4_6_release/'
                'apis/3.0/%s' % section)
    # used in the requests module for authentication for the API
    header = {'Authorization': 'Basic ' + b64_clientid(client_id)}
    # variables passed in URL as ?pagesSize=100&page=....&orderBy=....
    payload = {'pageSize': '100', 'page': page, 'orderBy': 'id desc',
               'conditions': 'closedBy = null and board/id = 1'}
    # HTTP GET using requests module to grab CW data via REST API
    r = requests.get(API_base + '/' + subsection,
                     params=payload, headers=header)
    return r.json()


def parse_JSON():
    """
    Parse the JSON data from cwJSON and format it into usable data. Assign the
    ticket numbers and statuses to dictionaries so we can find the length of
    those dictionaries later.
    """
    # need to call all the global variables to be used again later
    global parsed_json, status_fu, status_qa, status_re, status_et
    global status_dis, status_wr, status_sch, status_pro, status_new
    # used as local variables for just this function
    page = 1
    index = 0

    # this will loop through each page of JSON until it finds and empty page
    # if the page is not empty, add it to the parsed_json global variable
    while True:
        if cwJSON('service', 'tickets', page) == []:
            break
        else:
            parsed_json += cwJSON('service', 'tickets', page)
            print(cwJSON('service', 'tickets', page))
            page += 1
            
    # check if index is bigger than the length of parsed_json, while it's not
    # find all the statuses from tickets and put them in global dictionary
    while index < len(parsed_json):
        pisid = parsed_json[index]['status']['name']

        if pisid == 'Follow Up':
            status_fu[parsed_json[index]['id']] = pisid
        elif pisid == 'QA Call':
            status_qa[parsed_json[index]['id']] = pisid
        elif pisid == 'Re-Opened':
            status_re[parsed_json[index]['id']] = pisid
        elif pisid == 'Waiting on Response':
            status_wr[parsed_json[index]['id']] = pisid
        elif pisid == 'Enter Time':
            status_et[parsed_json[index]['id']] = pisid
        elif pisid == 'Discuss':
            status_dis[parsed_json[index]['id']] = pisid
        elif pisid == 'Scheduled':
            status_sch[parsed_json[index]['id']] = pisid
        elif pisid == 'In Progress':
            status_pro[parsed_json[index]['id']] = pisid
        elif pisid == 'New' or pisid == 'New (email)':
            status_new[parsed_json[index]['id']] = pisid
        index += 1


def outputFile():
    """
    Output the length of each dictionary to the fields.php file to be used in
    the status board. Each dictionary length will be used as a variable to
    display for statistics.
    """
    # v2 script - Output the data to an .ini file, easier to read and then
    # update instead of overwriting each time. Use parse_ini_file in PHP to
    # read.

    d = datetime.date.today()
    date = '%s' % str(d.year) + str(d.month)

    outputFile = configparser.ConfigParser()
    outputFile.read('fields.ini')

    # record the highest ticket count for the month for historical records
    # check to see if the month exists, if it does see if it's at the all time
    # high or not and set it, it not then make the value.
    if outputFile.has_option('Historical', '%s' % date):
        if int(outputFile['Historical']['%s' % date]) < len(parsed_json):
            outputFile['Historical']['%s' % date] = '%s' % len(parsed_json)
    else:
        outputFile.add_section('Historical')
        outputFile['Historical']['%s' % date] = '%s' % len(parsed_json)

    # output all the current ticket status counts to file
    outputFile['Ticket Status']['ticket_count'] = '%s' % len(parsed_json)
    outputFile['Ticket Status']['status_new'] = '%s' % len(status_new)
    outputFile['Ticket Status']['status_re'] = '%s' % len(status_re)
    outputFile['Ticket Status']['status_dis'] = '%s' % len(status_dis)
    outputFile['Ticket Status']['status_sch'] = '%s' % len(status_sch)
    outputFile['Ticket Status']['status_pro'] = '%s' % len(status_pro)
    outputFile['Ticket Status']['status_fu'] = '%s' % len(status_fu)
    outputFile['Ticket Status']['status_wr'] = '%s' % len(status_wr)
    outputFile['Ticket Status']['status_et'] = '%s' % len(status_et)
    outputFile['Ticket Status']['status_qa'] = '%s' % len(status_qa)

    with open('fields.ini', 'w') as output:
        outputFile.write(output)

# run the actual functions to grab data, parse it, and output it
parse_JSON()
outputFile()
