# apiCW.py - Connectwise API Query and JSON Parsing
I'm working on this script to connect to the Connectwise REST API and download
data for reporting.  It could probably be used for POST instead of GET, but I am
focusing on reports for work.  I had the initial version of the code downloading
ticket data and counting the status types of open tickets for a PHP dashboard
display we have in our call center, but I wanted a more generic data download
script that I can use for time and scheduling reports.  So I've re-written the
code as a module/class now, it is still being worked on when I have time but I
am much happier with the current iteration of the code.