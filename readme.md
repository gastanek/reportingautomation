<h3>My Reporting Automation</h3>

<u>Background</u> I need to produce a report each week based on metrics and datapoints that are collected in another system.  In my case the system is Salesforce (but I am not able to get the priviliges to use SFDC api's to fetch report data).  The data is reported in a Google Spreadsheet.  I used Google app script to automate my reporting.

Basics:
- the master report is done in a Google Spreadsheet
- report data is automatically emailed to me from salesforce
- i have a private Google spreadsheet that hosts this code as Google script
- the script collects data from the email reports, other inputs, stages it, and posts it to the master report


 How it works is simple:

<ol>
<li>SFDC is set up to autmatically email me my reports</li>
<li>Report emails are filtered and have a specific reporting label applied to them as they arrive</li>
<li>The app is exposed to Slack, with a slack interaction created, that allows me to post updates to the report from Slack</li>
<li>At a scheduled time each week, the main process is called to produce my report</li>
<li>The report data is collected from my email by finding the threads with the specific label</li>
<li>The report data is stripped of sfdc formatting and raw text added into tabs</li>
<li>The script then collects and stages the updates in a specific anchor sheet that is exactly the same as the master report</li>
<li>It runs through a sanity check just to be sure things seem ok</li>
<li>Once confirmed to be sane, it takes the clean report from the anchor sheet and coppies it directly to the master report</li>
<li>It uses specific text it expects to find plus a specific layout of the report fields to know where to paste the contents from the anchor sheet to the master report.</li>
<li>Afterwards it takes snapshots for historical pictures and then emails me a copy so I can see what was posted.</li>
<li>It then cleans everything up, including my email, so the next week new reporting emails are collected and things start fresh</li>
</ol>


Notes:
This is all very fragile because it depends significantly on an expected layout.  In my private report that is easy to control.  In the Master report that isn't so easy to control.  It's a downside byproduct of using a tool like Spreadsheets for a weekly reporting process.  It's an easy downside to overlook if you set it up with the expectation that human beings will be manually filling in the report each week and thus able to apply their own intelligence to the process.
