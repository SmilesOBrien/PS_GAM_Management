# PS_GAM_Management
A PSE (Plain Stupid English) interface using PowerShell to manage Google Workspace accounts and Chrome Devices

I wrote this script as a way to easily manage users and devices in Google Workspace for the school I work at. This was inspired by various posts I saw where people leveraged PowerShell to interact with GAM. I love GAM and what it does, but I didn't want to have to memorize commands and syntaxes or have to constantly go back to a cheatsheet. I just wanted an easy way to put in a device Asset ID and let it rip.

This script has the following pre-requisites:
1. Google Apps Manager (GAM)
2. The PowerShell Module "ImportExcel"
3. An Excel document in which you input your data

Please note: this was written specifically with my environment in mind, please edit the script as appropriate for your environment

I'm not a coder / scripter, just a tech who wanted a better way to leverage his tools. If you have any suggestions please let me know! I can't guarantee a swift turn-around but I'm happy to make this tool better!

HOW TO USE:

1. Edit the variables as appropriate for your environment
2. Edit any of the Write-Host statements as appropriate for your environment, and the corresponding language in the script below

NOTE: The environment this was written for uses the Asset ID field in Google Workspace. You may need to edit this script to work better in your environment off of a different descriptor. Any suggestions on making this more interoperable would be welcome.
