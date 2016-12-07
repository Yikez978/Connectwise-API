#! python3
# -*- coding: utf-8 -*-

"""grab_data.py.

Use the apiCW.py module to pull data from Connectwise.  No formatting yet.

This is strictly to test the different data queries before creating reports.
"""

import apiCW

# Dictionary of api login info
cwInfo = {"companyName": "COMPANY_ID",
          "pubKey": "COMPANY_PUBKEY",
          "privKey": "COMPANY_PRIVKEY"}

# initialize apiCW class, use cwInfo as keyword arguments
newData = apiCW.apiCW(**cwInfo)

# Service Ticket entries, page 1, Closed, on Service board
query1 = {"sect": "service", "subsect": "tickets",
          "page": "1",
          "condit": "closedFlag=True and board/name=\"Service\""}

# Open Tickets on Internal Board
query2 = {"sect": "service", "subsect": "tickets",
          "page": "1",
          "condit": "board/name=\"Internal\" and closedFlag=False"}

# Service Ticket 77740 Time Entries
query3 = {"sect": "service", "subsect": "tickets/77740/timeentries",
          "page": "1",
          "condit": ""}

# How many total Open Tickets on Service board?
query4 = {"sect": "service", "subsect": "tickets/count",
          "page": "1",
          "condit": "closedFlag=False and board/name=\"Service\""}

# Ticket Search on Service board, using search terms
# ToDo change get to post, make new kwarg
query5 = {"sect": "service", "subsect": "tickets/search",
          "page": "1",
          "condit": "closedFlag=True and board/name=\"Service\""}

#print(newData.dlData(**query5))


# Find all pages of Open Tickets, print each ticket as json
query = {"sect": "service", "subsect": "tickets", "page": 1,
         "condit": "closedFlag=False and board/name=\"Service\""}

while True:
    data = newData.dlData(**query)

    if data != []:
        for ticket in data:
            print(ticket)  # need to serve this to formatting functions
        #print(data)
        query['page'] += 1
    else:
        break
