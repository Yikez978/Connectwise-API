#! python3
# -*- coding: utf-8 -*-

"""apiCW.py.

Pull data from Connectwise API, generate reports.
"""
import base64
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
        """apiCW initialization.

        Assigns variables needed in methods.
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


class reportCW(apiCW):
    """reportCW class.

    Reporting class for Connectwise data.  Uses apiCW class as base to grab
    and parse JSON data into usable formats.  This is used to generate reports
    based on already downloaded/parsed data.
    """

    def __init__(self, sect, subsect, startDate, endDate):
        """reportCW initialization.

        Assigns variables needed in methods and calls
        the parseJSON method from apiCW to get the parsed dataset.
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

            worksheet1.set_column(0, col + 1, 11)  # expand A-F for header

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
    report = reportCW('time', 'entries', '2016-08-14', '2016-08-18')
    report.reportTimePerTech()
