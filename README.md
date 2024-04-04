# PS_GAM_Management
A PSE (Plain Stupid English) interface using PowerShell to manage Google Workspace accounts and Chrome Devices

I wrote this script as a way to easily manage users and devices in Google Workspace for the school I work at. This was inspired by various posts I saw where people leveraged PowerShell to interact with GAM. I love GAM and what it does, but I didn't want to have to memorize commands and syntaxes or have to constantly go back to a cheatsheet. I just wanted an easy way to put in a device Asset ID and let it rip.

This script has the following pre-requisites:
1. Google Apps Manager (GAM)
2. The PowerShell Module "ImportExcel"
3. An Excel document in which you input your data

Please note: this was written specifically with my environment in mind, please edit the script as appropriate for your environment

I'm not a coder / scripter, just a tech who wanted a better tool. If you have any suggestions please let me know!

HOW TO USE:

1. Edit the variables for $ExcelDB, $Worksheet, and $Column as appropriate for your environment
2. Edit any of the Write-Host statements as appropriate for your environment, and the corresponding language in the script below

PLANNED CHANGES:
1. Adding language use as variables for easier editing (for e.g. $lang1 = "Input Asset ID without leading 0" being set at the top of the script, and called where approrpiate, so if edit is needed you only need to edit in one place)
2. Adding nested menues, to control how long the Main Menu of the script is at any one time (for e.g. "Are you working on a User or a Device" etc)
