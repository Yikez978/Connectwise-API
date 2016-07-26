# -----------------------------------------------------------------------------
# Script: CWdata.ps1
# Author: Dustin Rowland
# Date: 01/25/2016
# Keywords: Connectwise, REST API
# Comments: Connect to REST API for Connectwise, download JSON data
# -----------------------------------------------------------------------------
# Client specific variables
$companyName = '<get from CW API>'
$pubKey = '<get from CW API>'
$privKey = '<get from CW API>'
# -----------------------------------------------------------------------------
# Convert the client variables into usable base64 encoded key, add to header info
$clientID = "$companyName+${pubKey}:$privKey" # same as ($companyName + '+' + $pubKey + ':' + $privKey)
$utf8_clientID = [System.Text.Encoding]::UTF8.GetBytes($clientID)
$b64_clientID = [System.Convert]::ToBase64String($utf8_clientID)
# -----------------------------------------------------------------------------
Function cwData($section, $subsection, $page) {

    $params = 'conditions=closedBy = null and board/id = 1&orderBy=id desc&pageSize=100&page='
    $base_url = 'https://api-na.myconnectwise.net/v4_6_release/apis/3.0/{0}/{1}?{2}{3}' -f $section,$subsection,$params,$page

    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", 'Basic ' + $b64_clientID)
    $cw_data = Invoke-RestMethod -Uri $base_url -Headers $headers

    $cw_data
}
# -----------------------------------------------------------------------------
