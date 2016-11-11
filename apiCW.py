#! python3
# -*- coding: utf-8 -*-

"""apiCW.py.

Pull data from Connectwise REST API.
"""
import base64
import requests    # requires PIP installation

__author__ = 'Dustin Rowland'
__email__ = 'lee.rowland@gmail.com'
__copyright__ = '2016, Dustin Rowland'
__status__ = 'Development'


class apiCW():
    """apiCW class.

    Primary utility class to connect to Connectwise servers and run queries
    against REST API.  Referenced below in reportCW class.
    """

    def __init__(self, companyName="changeme", pubKey="changeme2",
                 privKey="changeme3"):
        """__init__ method.

        Used to initialize the class, assign values to variables, run
        functions, etc.
        """
        self.companyName = companyName
        self.pubKey = pubKey
        self.privKey = privKey

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

    def dlData(self, sect="section", subsect="subsection", page="page#",
               condit="conditions"):
        """dlData method.

        Download the requested data using the Connectwise API. This will pull
        a JSON response for parsing later. Uses the b64_clientid function in
        the header variable.
        """
        API_base = ('https://api-na.myconnectwise.net/v4_6_release/'
                    'apis/3.0/{}'.format(sect))
        header = {'Authorization': 'Basic {}'.format(self.b64ClientID()),
                  'cache-control': 'no-cache'}
        payload = {'pageSize': '100', 'page': page, 'orderBy': 'id desc',
                   'conditions': condit}
        r = requests.get(API_base + '/' + subsect, params=payload,
                         headers=header)
        return r.json()


if __name__ == "__main__":
    print("You have called apiCW directly. Please use a launcher script.")
