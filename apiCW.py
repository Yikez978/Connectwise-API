#! python3
# -*- coding: utf-8 -*-

"""apiCW.py.

Pull data from Connectwise API, generate reports.
"""
import base64
import datetime
import requests  # requires PIP installation
import xlsxwriter  # requires PIP installation

__author__ = 'Dustin Rowland'
__email__ = 'lee.rowland@gmail.com'
__copyright__ = '2016, Dustin Rowland'
__status__ = 'Development'

# Look for ToDo comments for changes/projects list


class apiCW():
    def __init__(self):
        # self.companyName = '<CW API CompanyName>'
        # self.pubKey = '<CW API public key>'
        # self.privKey = '<CW API private key>'
        self.companyName = 'protech'
        self.pubKey = 'NFQIspb9XtiIlpT5'
        self.privKey = 'WyGglWJQxhWswUAS'

    def b64ClientID(self):
        """
        Uses the defined companyName, pubKey, and privKey to generate a
        client_id. Encodes the client_id at utf-8, instead of ASCII, then
        converts it to base64, then decode as utf-8 to remove byte 'b' prefix
        and \n newline.
        """
        client_id = '{}+{}:{}'.format(self.companyName, self.pubKey,
                                      self.privKey)
        cid_encode = client_id.encode('utf-8')
        cid_base64 = base64.b64encode(cid_encode)
        cid_decode = cid_base64.decode('utf-8')

        return cid_decode

    def formatDate(self, datestamp):
        return datetime.datetime.strptime(datestamp[0:16], '%Y-%m-%dT%H:%M')

    def queryAPI(self, section, subsection, page):
        """
        Query the Connectwise API for a specified request. This will pull down
        a JSON response for parsing later. Uses the b64_clientid function in
        the header variable.
        """
        API_base = ('https://api-na.myconnectwise.net/v4_6_release/'
                    'apis/3.0/{}'.format(section))
        header = {'Authorization': 'Basic {}'.format(self.b64ClientID())}
        payload = {'pageSize': '100', 'page': page, 'orderBy': 'id desc',
                   'conditions': ''}
        r = requests.get(API_base + '/' + subsection, params=payload,
                         headers=header)
        return r.json()

    def parseJSON(self, section, subsection, startDate, endDate):
        """
        Call queryAPI method to pull dataset out of online system. Will parse
        JSON for data between date ranges, pull out key information.
        """
        timeEntries = []
        startDate = datetime.datetime.strptime(startDate, '%Y-%m-%d')
        endDate = datetime.datetime.strptime(endDate, '%Y-%m-%d')
        page = 1

        while page < 11:  # ToDo: need to verify 10 pages of tickets is enough
            for i in self.queryAPI(section, subsection, page):
                # Verify entry timeStart is between startDate and endDate
                if startDate <= self.formatDate(i['timeStart']) <= endDate:
                    """ ToDo
                    This needs to create a nested dictionary of ticket
                    information based on enteredBy. Then we can pass the
                    dictionary to the reports for sumation.
                    """
                    timeEntries.append((i['enteredBy'], i['billableOption'],
                                        i['hoursBilled'], i['chargeToId'],
                                        i['workType']['name'],
                                        self.formatDate(i['timeStart']),
                                        self.formatDate(i['timeEnd']),
                                        i['company']['name']))
            page += 1
        return timeEntries  # return list with each entry as a tuple


class reportCW(apiCW):
    def __init__(self):
        apiCW.__init__(self)

    def reportTicketsPerCompany(self, section, subsection, startDate, endDate):
        """
        .
        """
        pass

    def reportTimePerTech(self, section, subsection, startDate, endDate):
        """
        Look for technician names in JSON data, count number of hours billed,
        then output to Excel file with chart.
        """
        timestamp = datetime.datetime.now()
        now = timestamp.strftime('%Y-%m-%d')
        reportData = self.parseJSON(section, subsection, startDate, endDate)

        """ ToDo
        Use data from parseJSON to create new variables based on dictionary
        keys, use those variables to finish report.  Do NOT hard-code techs
        or names for techs, if we add/change employees this section has to be
        modified everytime.  Make this more general, not specific.

        for tech in reportData:
            for ticket in tech:
                time += int(tech[ticket]['hoursBilled'])
            tech['timeSum'] = time
        """
        drowland = []
        rkoutz = []
        dyeager = []
        kmasoner = []
        drTime = rkTime = dyTime = kmTime = 0

        for entry in reportData:
            if entry[0] == 'rkoutz':
                rkoutz.append(entry)
            elif entry[0] == 'DYeager':
                dyeager.append(entry)
            elif entry[0] == 'kmasoner':
                kmasoner.append(entry)
            elif entry[0] == 'DRowland':
                drowland.append(entry)
            else:
                continue

        for time in rkoutz:
            rkTime += float(time[2])
        for time in dyeager:
            dyTime += float(time[2])
        for time in kmasoner:
            kmTime += float(time[2])
        for time in drowland:
            drTime += float(time[2])

        """ ToDo
        Once tech time is figured out, re-write report to use the new variable
        techs to generate report/chart.
        """"

        with xlsxwriter.Workbook('Report.xlsx') as workbook:
            workbook.set_properties({
                'title': 'Time Entries Report',
                'subject': 'Time Entries Report per Tech',
                'comments': 'Generated at ' + now
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

            worksheet1 = workbook.add_worksheet('Time Report')
            worksheet1.set_landscape()
            worksheet1.set_paper(1)         # 8x11 paper
            worksheet1.fit_to_pages(1, 1)   # scale down to fit on 1 paper
            worksheet1.set_margins(left=0.7, right=0.7, top=0.5, bottom=0.25)
            worksheet1.set_row(40)
            worksheet1.merge_range('A1:I1',
                                   'Time Entry Report for {} - {}'.format(
                                                           startDate, endDate),
                                   header_format)
            worksheet1.set_column(0, 0, 11)
            worksheet1.write('A2', 'D. Rowland', colhead_format)
            worksheet1.set_column(1, 1, 11)
            worksheet1.write('B2', 'R. Koutz', colhead_format)
            worksheet1.set_column(2, 2, 11)
            worksheet1.write('C2', 'K. Masoner', colhead_format)
            worksheet1.set_column(3, 3, 11)
            worksheet1.write('D2', 'D. Yeager', colhead_format)

            worksheet1.write(2, 0, drTime)
            worksheet1.write(2, 1, rkTime)
            worksheet1.write(2, 2, kmTime)
            worksheet1.write(2, 3, dyTime)

            chart = workbook.add_chart({'type': 'column'})
            chart.set_title({'name': 'Hours per Tech over Period'})
            chart.add_series({
                                'categories': ['Time Report', 1, 0, 1, 3],
                                'values': ['Time Report', 2, 0, 2, 3],
                                'data_labels': {'value': True}
                            })
            worksheet1.insert_chart('A4', chart)


if __name__ == "__main__":
    report = reportCW()
    report.reportTimePerTech('time', 'entries', '2016-08-14', '2016-08-18')
