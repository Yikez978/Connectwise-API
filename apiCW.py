#! python3
# -*- coding: utf-8 -*-

"""apiCW.py.

Pull data from Connectwise API, generate reports.
"""
import base64
import sqlite3
import datetime
import requests    # requires PIP installation
import xlsxwriter  # requires PIP installation

__author__ = 'Dustin Rowland'
__email__ = 'lee.rowland@gmail.com'
__copyright__ = '2016, Dustin Rowland'
__status__ = 'Development'


class apiCW():
    """apiCW class.

    Primary utility class to connect to Connectwise servers and run queries
    against REST API.  Referenced below in reportCW class.
    """

    def __init__(self):
        """__init__ method.

        Used to initialize the class, assign values to variables, run
        functions, etc.
        """
        self.companyName = '<CW API CompanyName>'
        self.pubKey = '<CW API public key>'
        self.privKey = '<CW API private key>'

    def b64ClientID(self):
        r"""b64ClientID method.

        Use the defined companyName, pubKey, and privKey to generate a
        client_id. Encode the client_id at utf-8, instead of ASCII, then
        convert it to base64, then decode as utf-8 to remove byte 'b' prefix
        and \n newline.
        """
        client_id = '{}+{}:{}'.format(self.companyName, self.pubKey,
                                      self.privKey)
        cid_encode = client_id.encode('utf-8')
        cid_base64 = base64.b64encode(cid_encode)
        cid_decode = cid_base64.decode('utf-8')

        return cid_decode

    def formatDate(self, datestamp):
        """formatDate method.

        Quick method to properly format the datestamp from the REST API into
        Python formats.
        """
        return datetime.datetime.strptime(datestamp[0:16], '%Y-%m-%dT%H:%M')

    def queryAPI(self, sect, subsect, page):
        """queryAPI method.

        Query the Connectwise API for a specified request. This will pull down
        a JSON response for parsing later. Uses the b64_clientid function in
        the header variable.
        """
        API_base = ('https://api-na.myconnectwise.net/v4_6_release/'
                    'apis/3.0/{}'.format(sect))
        header = {'Authorization': 'Basic {}'.format(self.b64ClientID())}
        payload = {'pageSize': '100', 'page': page, 'orderBy': 'id desc',
                   'conditions': ''}
        r = requests.get(API_base + '/' + subsect, params=payload,
                         headers=header)
        return r.json()

    # ToDo: Phase out after making xlsxCW use sqlCW for data
    def parseJSON(self, sect, subsect, startDate, endDate):
        """parseJSON method.

        Call queryAPI method to pull dataset out of online system. Will parse
        JSON for data between date ranges, pull out key information.
        """
        techEntries = {}
        startDate = datetime.datetime.strptime(startDate, '%Y-%m-%d')
        endDate = datetime.datetime.strptime(endDate, '%Y-%m-%d')
        page = 1

        while page < 11:  # ToDo: need to verify 10 pages of tickets is enough
            for i in self.queryAPI(sect, subsect, page):
                # Verify entry timeStart is between startDate and endDate
                if startDate <= self.formatDate(i['timeStart']) <= endDate:
                    ieB = i['enteredBy']
                    icTI = i['chargeToId']
                    ihB = i['hoursBilled']
                    ibO = i['billableOption']
                    iwn = i['workType']['name']
                    itS = i['timeStart']
                    itE = i['timeEnd']
                    icn = i['company']['name']

                    if ieB not in techEntries:
                        techEntries[ieB] = {}
                    if icTI not in techEntries[ieB]:
                        techEntries[ieB][icTI] = {}

                    techEntries[ieB][icTI]['hoursBilled'] = ihB
                    techEntries[ieB][icTI]['billableOption'] = ibO
                    techEntries[ieB][icTI]['workType'] = iwn
                    techEntries[ieB][icTI]['timeStart'] = self.formatDate(itS)
                    techEntries[ieB][icTI]['timeEnd'] = self.formatDate(itE)
                    techEntries[ieB][icTI]['companyName'] = icn
            page += 1
        return techEntries


class sqlCW(apiCW):
    """sqlCW class.

    Takes the information from cwAPI queryAPI and sends the JSON output into a
    SQLite database.  This will be used for longterm reporting, tracking tech
    time entries over a couple years.
    """
    """ToDo.
    [ ] Write new parseJSON function to transfer JSON into SQL insert
    [x] Build SQLite creation (done in pysswdmgr, make a utility module?)
        [x] Build tech table, if it doesn't exist - id, name, cwname
        [x] Build ticket table, if it doesn't exist - entryID, tech.id, entry
            type (ticket, activity, etc), company name, time billed, time type
            (billed, not billed, do not charge)
    [ ] Create SQLite insert function
        [ ] Verify if entryID exists
    [ ] Create views or queries for weekly, monthly, yearly reports
    """
    def __init__(self, dbname='cwTime.db'):
        """__init__ method.

        Used to initialize the class, assign values to variables, run
        functions, etc.
        """
        self.db = dbname
        self.conn = sqlite3.connect(self.db)
        self.cur = self.conn.cursor()

    def queryDB(self, arg):
        """queryDB method. Used by functions to manipulate the SQLite DB."""
        self.cur.execute(arg)
        self.conn.commit()
        return self.cur

    def table_verify(self):
        """table_verify method.

        Query to the DB, create tables if they do not exist.
        """
        # name = JSON member name
        # techname = JSON member identifier
        # techid = JSON member id
        self.queryDB("""CREATE TABLE IF NOT EXISTS
                        tech ("index" INTEGER PRIMARY KEY AUTOINCREMENT
                        UNIQUE, name TEXT, techname TEXT, techid INT)""")
        # name = JSON company name
        # clientname = JSON company identifier
        # clientid = JSON company id
        self.queryDB("""CREATE TABLE IF NOT EXISTS
                        client ("index" INTEGER PRIMARY KEY AUTOINCREMENT
                        UNIQUE, name TEXT, clientname TEXT, clientid INT)""")
        # techid = JSON member id
        # clientid = JSON company id
        # chargeid = JSON chargeToID
        # chargetype = JSON chargeToType
        # billtype = JSON billableOption
        # worktype = JSON workType name
        # timestart = JSON timeStart (use apiCW formatDate)
        # timeend = JSON timeEnd (use apiCW formatDate)
        # timecharge = JSON hourBilled
        self.queryDB("""CREATE TABLE IF NOT EXISTS
                        ticket ("index" INTEGER PRIMARY KEY AUTOINCREMENT
                        UNIQUE, techid INT, clientid INT, chargeid TEXT,
                        chargetype TEXT, billtype TEXT, worktype TEXT,
                        timestart TEXT, timeend TEXT, timecharge INT)""")


class xlsxCW(apiCW):  # ToDo: rewrite to pull data from sqlCW, not apiCW
    """xlsxCW class.

    Reporting class for Connectwise data.  Uses apiCW class as base to grab
    and parse JSON data into usable formats.  This is used to generate reports
    based on already downloaded/parsed data.
    """

    def __init__(self, sect, subsect, startDate, endDate):
        """__init__ method.

        Used to initialize the class, assign values to variables, run
        functions, etc.  Calls the parseJSON method from apiCW to get the
        parsed dataset.
        """
        apiCW.__init__(self)
        self.timestamp = datetime.datetime.now()
        self.now = self.timestamp.strftime('%Y-%m-%d')
        self.repData = self.parseJSON(sect, subsect, startDate, endDate)
        self.startDate = startDate
        self.endDate = endDate

    def reportTicketsPerCompany(self):
        """reportTicketsPerCompany method.

        Look at each company name listed in self.repData, count number of
        tickets opened over period, then output to Excel file with chart.
        """
        pass

    def reportTimePerTech(self):
        """reportTimePerTech method.

        Look for technician names in self.reportData, count number of hours
        billed, then output to Excel file with chart.
        """
        with xlsxwriter.Workbook('TimePerTech_Report.xlsx') as workbook:
            workbook.set_properties({
                'title': 'Time Entries Report',
                'subject': 'Time Entries Report per Tech',
                'comments': 'Generated at ' + self.now
            })

            header_format = workbook.add_format({'align': 'center',
                                                 'align': 'vcenter',
                                                 'bg_color': '#99ff99',
                                                 'font_color': '#e68a00',
                                                 'font_size': '24',
                                                 'bold': True,
                                                 'bottom': 1,
                                                 'num_format': 'yyyy-mm-dd'})
            colhead_format = workbook.add_format({'bg_color': '#004d00',
                                                  'font_color': '#ffd699',
                                                  'bottom': 1})

            worksheet1 = workbook.add_worksheet('TimeReport')
            worksheet1.set_landscape()
            worksheet1.set_paper(1)         # 8x11 paper
            worksheet1.fit_to_pages(1, 1)   # scale down to fit on 1 paper
            worksheet1.set_margins(left=0.7, right=0.7, top=0.5, bottom=0.25)
            worksheet1.set_row(40)
            worksheet1.merge_range('A1:H1',
                                   'Time Entry Report for {} - {}'.format(
                                                           self.startDate,
                                                           self.endDate),
                                   header_format)

            row = 1
            col = 0
            for tech in self.repData:
                tTime = 0

                for ticket in self.repData[tech]:
                    tTime += float(self.repData[tech][ticket]['hoursBilled'])
                self.repData[tech]['totalTime'] = tTime

                worksheet1.write_string(row, col, tech, colhead_format)
                worksheet1.write_number(row + 1, col, tTime)
                col += 1

            worksheet1.set_column(0, col + 1, 11)  # Expand A-F for header

            chart = workbook.add_chart({'type': 'column'})
            chart.set_size({'width': 620, 'height': 300})
            chart.set_title({'name': 'Hours per Tech over Period'})
            chart.set_legend({'none': True})
            chart.add_series({
                                'categories': ['TimeReport', 1, 0, 1, col - 1],
                                'values': ['TimeReport', 2, 0, 2, col - 1],
                                'data_labels': {'value': True}
                            })
            worksheet1.insert_chart('A4', chart)


if __name__ == "__main__":
    print("You have called apiCW directly. Please use a launcher script.")
