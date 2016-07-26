# -----------------------------------------------------------------------------
# Script: ticket_data.ps1
# Author: Dustin Rowland
# Date: 01/25/2016
# Keywords: Connectwise, REST API
# Comments: Pull information using CWData.ps1, parse ticket information
# -----------------------------------------------------------------------------
# Include CWData.ps1 module for Connectwise REST API
. 'P:\Connectwise API\CWdata.ps1'
# -----------------------------------------------------------------------------
# Script variables
$page = 1  #used for cwData
$index = 0
$status_fu = New-Object System.Collections.ArrayList
$status_re = New-Object System.Collections.ArrayList
$status_et = New-Object System.Collections.ArrayList
$status_wr = New-Object System.Collections.ArrayList
$status_dis = New-Object System.Collections.ArrayList
$status_sch = New-Object System.Collections.ArrayList
$status_pro = New-Object System.Collections.ArrayList
$status_new = New-Object System.Collections.ArrayList

while($true) {

    if ((cwData -section service -subsection tickets -page $page) -eq $null) { 
        break
    }
    else {
        $cwJSON += cwData -section service -subsection tickets -page $page
        $page++
    }

}

while($index -lt $cwJSON.Count) {
    # during this you will see lots of numbers, it's the arrays updating index with new items
    if ($cwJSON[$index].status.name -eq 'In Progress') {$status_pro.Add($cwJSON[$index].id)}
    elseif ($cwJSON[$index].status.name -eq 'Follow Up') {$status_fu.Add($cwJSON[$index].id)}
    elseif ($cwJSON[$index].status.name -eq 'Re-Opened') {$status_re.Add($cwJSON[$index].id)}
    elseif ($cwJSON[$index].status.name -eq 'Enter Time') {$status_et.Add($cwJSON[$index].id)}
    elseif ($cwJSON[$index].status.name -eq 'Waiting on Response') {$status_wr.Add($cwJSON[$index].id)}
    elseif ($cwJSON[$index].status.name -eq 'Discuss' -or $cwJSON[$index].status.name -eq 'Hibernation') {$status_dis.Add($cwJSON[$index].id)}
    elseif ($cwJSON[$index].status.name -eq 'Scheduled') {$status_sch.Add($cwJSON[$index].id)}
    elseif ($cwJSON[$index].status.name -like 'New*') {$status_new.Add($cwJSON[$index].id)}
    $index++
}

<#"In Progress: {0}" -f $status_pro.Count
"Follow Up: {0}" -f $status_fu.Count
"Re-Opened: {0}" -f $status_re.Count
"Enter Time: {0}" -f $status_et.Count
"Waiting on Response: {0}" -f $status_wr.Count
"Discuss/Hibernate: {0}" -f $status_dis.Count
"Scheduled: {0}" -f $status_sch.Count
"New: {0}" -f $status_new.Count#>

$jsonOut = @"
{
  "total": "$($cwJSON.Count)",
  "tickets": {
              "status_fu": "$($status_fu.Count)",
              "status_re": "$($status_re.Count)",
              "status_et": "$($status_et.Count)",
              "status_wr": "$($status_wr.Count)",
              "status_pro": "$($status_pro.Count)",
              "status_dis": "$($status_dis.Count)",
              "status_sch": "$($status_sch.Count)",
              "status_new": "$($status_new.Count)"
            }
}
"@

$jsonOut | Out-File -Encoding ascii "P:\Connectwise API\cwdata.json"
