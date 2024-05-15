# Planned Changes
 - Calling GAM commands as functions based on inputs, rather than as repeat code
 - Add support for Google Drive management
 - Update up get-csvandheaders code to skip null values in CSV (current this is handled when the fuction is called, isntead of in the function itself)
 - Add support for adding students to google classrooms

# start-PSGAM
A PSE (Plain Stupid English) interface using PowerShell to manage Google Workspace accounts and Chrome Devices

I wrote this script as a way to easily manage users and devices in Google Workspace for the school I work at. This was inspired by various posts I saw where people leveraged PowerShell to interact with GAM. I love GAM and what it does, but I didn't want to have to memorize commands and syntaxes or have to constantly go back to a cheatsheet. I just wanted an easy way to put in a device Asset ID and let it rip.

This script requires Google Apps Manager (GAM) in order to function.

Please note: this was written specifically with my environment in mind, please edit the script as appropriate for your environment

I'm not a coder / scripter, just a tech who wanted a better way to leverage his tools. If you have any suggestions please let me know! I can't guarantee a swift turn-around but I'm happy to make this tool better!

HOW TO USE:

- Call the script
- Select whether you are working with an Asset, and User, or wish to enter a manual GAM command
- Choose the option that is most applicable to your situation
- If making a bulk change, a file explorer window will prompt you to choose a CSV file.
- You can use one CSV for everything, the script will prompt you to select which header you are using.

NOTE: The environment this was written for uses the Asset ID field in Google Workspace. You may need to edit this script to work better in your environment off of a different descriptor. Any suggestions on making this more interoperable would be welcome.

# Current Capabilities
 - Find the serial number of single and bulk devices via the AssetID parameter
 - Move individual or bulk devices via the AssetID parameter
 - Wipe individual or bulk devices via the AssetID parameter
 - Enable / disable individual devices
 - Powerwash individual devices
 - Add or remove a user from a specified group
 - Delete emails from ALL inboxes via MessageID or Sender Address
 - Reset Google Password
 - Show Google Classrooms owned by specified teacher email address
 - Add teacher to individual or bulk courses and make owner
 - Run a manual command
