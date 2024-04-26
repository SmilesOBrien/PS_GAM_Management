    <#
    .SYNOPSIS
        A PSE (Plain Stupid English) interface using PowerShell to manage Google Workspace accounts and Chrome Devices

    .DESCRIPTION
        Without other parameters, $Header defaults to "AssetNum"
        More information can be found here https://github.com/smiles-obrien/PS_GAM_Management

    .PARAMETER csv
        Path to your CSV file

    .PARAMETER Header
        CSV header name, defaults to "AssetNum"

    .EXAMPLE
        PS> start-PSGAM -CSV "D:\My Drive\Sheets\Chromebook Management.xlsx"

    .EXAMPLE
        PS> start-PSGAM -CSV "D:\My Drive\Sheets\Chromebook Management.xlsx" -Header "CourseID"
    #>

    param(
        [Parameter(Mandatory)][string]$csv,
        [string]$Header = "AssetNum"
    )

    $lang1 = "Input Asset ID without leading 0"

    # Prompt for choice
    $Title = "============== GAM Powerhsell Interface ==============="
    $Prompt = "I need to edit:"
    $Choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Assets", "&Users", "&Manual GAM Command")
    $Default = 0

    # Prompt for the choice
    $Choice = $host.UI.PromptForChoice($Title, $Prompt, $Choices, $Default)

    # Action based on choice
    Switch($Choice) {
        0 {
            do {
                Write-Host "`n1. Retrieve Serial Number - Single Device" -ForegroundColor White
                Write-Host "2. Retrieve Serial Numbers - Bulk Devices" -ForegroundColor White
                Write-host "3. Move to New OU - Single Device" -ForegroundColor White
                Write-Host "4. Move to New OU - Bulk Devices" -ForegroundColor White
                Write-Host "5. Wipe Users From Device - Single Device" -ForegroundColor White
                Write-Host "6. Wipe Users From Device - Bulk Devices" -ForegroundColor White
                Write-Host "7. Disable Single Device" -ForegroundColor White
                Write-Host "8. Enable Single Device" -ForegroundColor White
                Write-Host "9. Powerwash Device - Single Device" -ForegroundColor White
                Write-Host "Q. Quit" -ForegroundColor Yellow
                [System.Environment]::NewLine
                $choice = Read-Host "Enter Choice"
                switch ($choice) {
            
                    '1'{ #Find single serial number based on inputted asset id, write output to host
            
                        [System.Environment]::NewLine
                    
                        $SingleAsset = Read-Host $lang1
            
                        $sn = @( gam info cros query:asset_id:$SingleAsset serialnumber asset_ID ) | ConvertFrom-String -delimiter "serialNumber: "
                            
                        $Results = [PSCustomObject]@{
                            assetNumber = $SingleAsset
                            serialNumber = $sn.p2
                        }
            
                        $Results | Out-host 
            
                        [System.Environment]::NewLine
            
                        Read-Host -Prompt "Press Enter to return to menu"
                    }
                    '2'{ #Find serial number of bulk devices, write output to host
            
                        [System.Environment]::NewLine
                        $AssetData = $AssetData = Import-csv -path $csv | Select-Object -ExpandProperty $Header
                        $( ForEach ( $asset in $AssetData ) {
            
                            $sn = @( gam info cros query:asset_id:$asset serialnumber asset_ID ) | ConvertFrom-String -delimiter "serialNumber: "
                            
                            $Results = [PSCustomObject]@{
                                assetNumber = $asset
                                serialNumber = $sn.p2
                            }
            
                            $Results
            
                        } ) | Out-host 
            
                        [System.Environment]::NewLine
            
                        Read-Host -Prompt "Press Enter to return to menu"
                    }
                    '3'{  #Using Asset tag as reference, move devices to desired Workspace OU, write output to host
            
                        [System.Environment]::NewLine
            
                        $SingleAsset = Read-Host $lang1
                        $OU = Read-Host "Enter Destination OU Path"
            
                        gam update cros query:asset_id:$SingleAsset ou "$OU" | Write-Host
            
                        [System.Environment]::NewLine
            
                        Read-Host -Prompt "Press Enter to return to menu"
            
                    }
                    '4'{ #Using Asset tag as reference, move devices to desired Workspace OU, write output to host
                    
                        $OU = Read-Host "Enter Destination OU Path"
                        $AssetData = Import-csv -path $csv | Select-Object -ExpandProperty $Header
                        ForEach ( $asset in $AssetData ) {
            
                            gam update cros query:asset_id:$asset ou "$OU" | Write-Host
                        }
            
                        [System.Environment]::NewLine
            
                        Read-Host -Prompt "Press Enter to return to menu"
                    }
                    '5' { #Using Asset tag as reference, wipe user accounts on single device, write output to host
                        
                        $SingleAsset = Read-Host $lang1
            
                        gam issuecommand cros query:asset_id:$SingleAsset command wipe_users doit | Write-Host
            
                        [System.Environment]::NewLine
            
                        Read-Host -Prompt "Press Enter to return to menu"
                    }
                    '6'{ #Using Asset tag as reference, wipe user accounts on bulk devices, write output to host
            
                        $AssetData = Import-csv -path $csv | Select-Object -ExpandProperty $Header
                        ForEach ( $asset in $AssetData ) {
            
                            [System.Environment]::NewLine
            
                            gam issuecommand cros query:asset_id:$asset command wipe_users doit | Write-Host
                        }
                        [System.Environment]::NewLine
            
                        Read-Host -Prompt "Press Enter to return to menu"
                    }
                    '7'{ #Using Asset tag as reference, Disable single device, write output to host
                        
                        [System.Environment]::NewLine
            
                        $SingleAsset = Read-Host $lang1
                        $sn = @( gam info cros query:asset_id:$SingleAsset serialnumber asset_ID ) | ConvertFrom-String -delimiter "serialNumber: "
                            
                        $Results = [PSCustomObject]@{
                            assetNumber = $SingleAsset
                            serialNumber = $sn.p2
                        }
            
                        gam update cros cros_sn $sn.p2 action Disable
                        [System.Environment]::NewLine
            
                        Read-Host -Prompt "Press Enter to return to menu"
                    }
                    '8'{ #Using Asset tag as reference, Enable single device, write output to host
                        
                        [System.Environment]::NewLine
            
                        $SingleAsset = Read-Host $lang1
                        $sn = @( gam info cros query:asset_id:$SingleAsset serialnumber asset_ID ) | ConvertFrom-String -delimiter "serialNumber: "
                            
                        $Results = [PSCustomObject]@{
                            assetNumber = $SingleAsset
                            serialNumber = $sn.p2
                        }
            
                        gam update cros cros_sn $sn.p2 action reenable
            
                        [System.Environment]::NewLine
            
                        Read-Host -Prompt "Press Enter to return to menu"
                    }
                    '9'{ #Using AssetID as reference, Powerwash single device

                    [System.Environment]::NewLine

                    $SingleAsset = Read-Host $lang1

                    gam issuecommand cros query:asset_id:$SingleAsset command remote_powerwash doit | Out-Host

                    [System.Environment]::NewLine

                    Read-Host -Prompt "Press Enter to return to menu"

                    }
                    'q'{ #Exit script

                    [System.Environment]::NewLine
                    Write-Output "Exiting Script"
                    Start-Sleep -seconds 2

                    Exit
                    }    
                } 
            } until ($choice -eq 'q')  
        }
        1 {
            do {
                Write-Host "1. Add User to Group - Single User" -ForegroundColor White
                Write-Host "2. Remove User From Group - Single User" -ForegroundColor White
                Write-Host "3. Delete Email from All Mailboxes - Message ID" -ForegroundColor White
                Write-Host "4. Delete Email from All Mailboxes - Sender Address" -ForegroundColor White
                Write-Host "5. Reset User Password - Single User (READ DISCLAIMER!)" -ForegroundColor White
                Write-Host "6. Show Google Classroom by User" -ForegroundColor White
                Write-Host "7. Update Google Classroom Owner - Bulk Classes" -ForegroundColor White
                Write-Host "8. Update Google Classroom Owner - Single Class" -ForegroundColor White
                Write-Host "Q. Quit" -ForegroundColor Yellow
                [System.Environment]::NewLine
                $choice = Read-Host "Enter Choice"
                switch ($choice) {

                    '1'{ #Adds specified user to specified group

                            $user = Read-Host "Enter User Email Address"
                            $Group = Read-host "Enter Group Email Address"
                            $type = Read-Host "Member or Manager?"

                            gam update group $group add $type $user | Out-Host

                            [System.Environment]::NewLine

                            Read-Host -Prompt "Press Enter to return to menu"

                    }
                    '2'{ #Removes specified user from specified group

                        $user = Read-Host "Enter User Email Address"
                        $Group = Read-host "Enter Group Email Address"

                        gam update group $group delete user $user | Out-Host

                        [System.Environment]::NewLine

                        Read-Host -Prompt "Press Enter to return to menu"
                    }
                    '3' { # Using MessageID as reference, delete email from all inboxes

                        [System.Environment]::NewLine
                        $email = Read-host "Input email address of reporter"
                        $MessageID = Read-Host "Input Message ID"
                        [System.Environment]::NewLine
                        Write-host "Checking all mailboxes for specified Message ID, and deleting. Operation will take approximately 3 minutes"

                            gam all users delete messages query rfc822msgid:$MessageID doit *> "$env:UserProfile\Documents\Phishing\'$email'_deleteEmail_$(get-date -Format yyyy-MM-dd_HHmm).txt"

                        Get-ChildItem -path $env:UserProfile\Documents\Phishing -filter "'$email'_deleteEmail*.txt" -recurse | Sort-Object CreationTime -Descending | Select-Object -First 1 | select-string -pattern 'Got 1 Messages'
                        [System.Environment]::NewLine
                        Read-Host -Prompt "Press Enter to return to menu"

                    }
                    '4' { # Using Sender Address as reference, delete email from all inboxes

                        [System.Environment]::NewLine
                        $email = Read-host "Input email address of reporter"
                        $sender = Read-Host "Input email address of sender"
                        [System.Environment]::NewLine
                        Write-host "Checking all mailboxes for specified Message ID, and deleting. Operation will take approximately 3 minutes"

                            gam all users delete messages query "from:$sender" doit *> "$env:UserProfile\Documents\Phishing\'$email'_deleteEmail_$(get-date -Format yyyy-MM-dd_HHmm).txt"

                        Get-ChildItem -path $env:UserProfile\Documents\Phishing -filter "'$email'_deleteEmail*.txt" -recurse | Sort-Object CreationTime -Descending | Select-Object -First 1 | select-string -pattern 'Got 1 Messages'
                        [System.Environment]::NewLine
                        Read-Host -Prompt "Press Enter to return to menu"

                    }
                    '5' { # Reset password via Google (READ DISCLAIMER!)

                        [System.Environment]::NewLine
                        
                        Write-Host "DISCLAIMER: PASSWORDS SYNC VIA ACTIVE DIRECTORY. ONLY RESET PASSWORD VIA GOOGLE IF THERE IS AN ISSUE WITH SYNC, OR AS EMERGENCY MEASURE" - -ForegroundColor DarkRed
                        [System.Environment]::NewLine
                        $Username = Read-Host "Input Email Address"
                        $Credential = Read-Host "Input New Password" -AsSecureString
                        [System.Environment]::NewLine

                        gam update user $username password $credential | Out-Host

                        [System.Environment]::NewLine
                        Read-Host -Prompt "Press Enter to return to menu"
                    }
                    '6'{ #Show Google Classroom by Teacher
                        $Teacher = read-host "Input Email Address"
                    
                        gam print courses teacher $Teacher fields name,id | Write-Host

                        [System.Environment]::NewLine

                        Read-Host -Prompt "Press Enter to return to menu"

                    }
                    '7'{ #Google Classroom bulk assign
                        $Teacher = read-host "Input Email Address"
                        $CourseData = Import-csv -path $csv | Select-Object -ExpandProperty $Header
                        ForEach ( $class in $CourseData ) {

                            gam course $class add teacher $Teacher | Write-Host
                            gam update course $class owner $Teacher | Write-Host
                        }

                        [System.Environment]::NewLine

                        Read-Host -Prompt "Press Enter to return to menu"

                    }
                    '8'{ #Google Classroom single assign
                        $Teacher = read-host "Input Email Address"
                        $CourseID = read-host "Input Course ID"

                            gam course $CourseID add teacher $Teacher | Write-Host
                            gam update course $CourseID owner $Teacher | Write-Host

                        [System.Environment]::NewLine

                        Read-Host -Prompt "Press Enter to return to menu"
                    }
                    'q'{ #Exit script

                        [System.Environment]::NewLine
                        Write-Output "Exiting Script"
                        Start-Sleep -seconds 2

                        Exit
                    }    
                } 
            } until ($choice -eq 'q')   
        }
        2 {
            $manual1 = Read-host "`nInput manual GAM command"
            Invoke-expression $manual1
                    
            do { $again = Read-Host "`nRun another command? (y/n)"
                        if ( $again -eq "y") {
                            $manual1 = Read-host "`nInput manual GAM command"
                            Invoke-expression $manual1
                        }
            } until ($again -eq "n")
        }
    }
