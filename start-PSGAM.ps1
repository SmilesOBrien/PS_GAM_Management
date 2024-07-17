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
        PS> start-PSGAM -CSV "D:\My Drive\Sheets\Chromebook Management.xlsx" -Header "ClassID"
    #>

# Below code from https://gist.github.com/joshooaj/63d1a88ea5bec2189442dd26c88de5b5
Add-Type -AssemblyName System.Windows.Forms
Add-Type @'
using System;
using System.Runtime.InteropServices;
public class WindowHelper {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
'@

# Below function adapted from https://gist.github.com/joshooaj/63d1a88ea5bec2189442dd26c88de5b5
function Show-OpenFileDialog {
    <#
    .SYNOPSIS
    Shows the Windows OpenFileDialog and returns the user-selected file path(s).
    .DESCRIPTION
    For detailed information on the available parameters, see the OpenFileDialog
    class documentation online at https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.openfiledialog?view=netframework-4.8.1
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter()]
        [bool]
        $AddExtension = $true,

        [Parameter()]
        [bool]
        $AutoUpgradeEnabled = $true,

        [Parameter()]
        [bool]
        $CheckFileExists = $true,

        [Parameter()]
        [bool]
        $CheckPathExists = $true,

        [Parameter()]
        [string]
        $DefaultExt,

        [Parameter()]
        [bool]
        $DereferenceLinks = $true,

        # Filter for specific file types. Example syntax: 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
        [Parameter()]
        [string]
        $Filter,

        [Parameter()]
        [string]
        $InitialDirectory,

        [Parameter()]
        [bool]
        $Multiselect,

        [Parameter()]
        [bool]
        $ReadOnlyChecked,

        [Parameter()]
        [bool]
        $RestoreDirectory,

        [Parameter()]
        [bool]
        $ShowHelp,

        [Parameter()]
        [bool]
        $ShowReadOnly,

        [Parameter()]
        [bool]
        $SupportMultiDottedExtensions,

        [Parameter()]
        [string]
        $Title,

        [Parameter()]
        [bool]
        $ValidateNames
    )

    process {
        $params = @{
            AddExtension                 = $AddExtension
            AutoUpgradeEnabled           = $AutoUpgradeEnabled
            CheckFileExists              = $CheckFileExists
            CheckPathExists              = $CheckPathExists
            DefaultExt                   = $DefaultExt
            DereferenceLinks             = $DereferenceLinks
            Filter                       = $Filter
            InitialDirectory             = $InitialDirectory
            Multiselect                  = $Multiselect
            ReadOnlyChecked              = $ReadOnlyChecked
            RestoreDirectory             = $RestoreDirectory
            ShowHelp                     = $ShowHelp
            ShowReadOnly                 = $ShowReadOnly
            SupportMultiDottedExtensions = $SupportMultiDottedExtensions
            Title                        = $Title
            ValidateNames                = $ValidateNames
        }

        [System.Windows.Forms.Form]$form = $null
        [System.Windows.Forms.OpenFileDialog]$dialog = $null
        try {
            $form = [System.Windows.Forms.Form]@{ TopMost = $true }
            $dialog = [System.Windows.Forms.OpenFileDialog]$params
            $CustomPlaces | ForEach-Object {
                if ($null -eq $_) {
                    return
                }
                if (($id = $_ -as [guid])) {
                    $dialog.CustomPlaces.Add($id)
                } else {
                    $dialog.CustomPlaces.Add($_)
                }
            }
            $null = [WindowHelper]::SetForegroundWindow($form.Handle)
            if ($dialog.ShowDialog($form) -eq 'OK') {
                if ($MultiSelect) {
                    $dialog.FileNames
                } else {
                    $dialog.FileName
                }
            } else {
                Write-Error -Message 'No file(s) selected.'
            }
        } finally {
            if ($dialog) {
                $dialog.Dispose()
            }
            if ($form) {
                $form.Dispose()
            }
        }
    }
}
function Get-CSVandHeaders {
    <#
    .SYNOPSIS
    Using the Show-OpenDialogue Box function above, import the selected CSV, and obtain the header information. 
    .DESCRIPTION
    Allows you to select a specific CSV header for use in bulk actions. For example, if you have a CSV with AssetNum, CourseID, UserID, you can specify which column of data you want to work with.
    #>
    # Import CSV
        $CSV = Show-OpenFileDialog
        $CSV
        $CSVData = Import-csv -path "$csv"

    # Obtain Headers
        $Headers = $CSVData[0].psobject.properties.name

    # Build a formatted string listing choices with numbers
        $choiceList = @{}
        for ($i = 0; $i -lt $Headers.Count; $i++) {
        $choiceList[($i + 1)] = "$($i + 1). $($Headers[$i])"
        }

    # Enumerate choices with numbers
        Write-Host "Available Choices:"
        for ($i = 0; $i -lt $Headers.Count; $i++) {
        Write-Host "  $($i + 1). $($Headers[$i])"
        }

    # Prompt the user for a choice using the array of strings
        $userChoiceIndex = Read-Host "Enter a choice (1 - $($Headers.Count))"

    # Validate user input (optional)
        if ($userChoiceIndex -lt 1 -or $userChoiceIndex -gt $headers.Count) {
        Write-Error "Invalid choice. Please enter a number between 1 and $($choices.Count)."
        }

    # Use the index to access the corresponding string from the choices array
        $userChoiceString = $headers[$userChoiceIndex - 1]

    # Write the chosen property name for debugging
        Write-Host "Selected Property Name: $userChoiceString"
        return $CSVData.$userChoiceString

    }

$lang1 = "Input Asset ID without leading 0"

# Prompt Info
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
            Write-host "9. Disable Bulk Devices" -ForegroundColor White
            Write-host "10. Enable Bulk Devices" -ForegroundColor White
            Write-Host "11. Powerwash Device - Single Device" -ForegroundColor White
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
                
                $Data = Get-CSVandHeaders | Where-Object { $_.PSObject.Properties.Value -ne '' } # This returns the CSV Path as the first line before the asset data, and skips empty cells. I plan to clean this in a future release.

                # Filter out the path CSV Path from $Data
                $FilteredData = $Data[1..($Data.Count)]
                
                $Results = $FilteredData | ForEach-Object {

                    # Get serial number using gam command (modify as needed)
                    $sn = @( gam info cros query:asset_id:$_ serialnumber asset_ID ) | ConvertFrom-String -delimiter "serialNumber: "

                    # Create Chart of Asset Numbers matched to Serial Numbers
                    [PSCustomObject]@{
                        assetNumber = $_
                        serialNumber = $sn.p2
                        }
                    }

                # Output results to host
                $Results | Out-host
 
                [System.Environment]::NewLine
    
                Read-Host -Prompt "Press Enter to return to menu"
            }
            '3'{ #Using Asset tag as reference, move single device to desired Workspace OU, write output to host
    
                [System.Environment]::NewLine
    
                $SingleAsset = Read-Host $lang1
                $OU = Read-Host "Enter Destination OU Path"
    
                gam update cros query:asset_id:$SingleAsset ou "$OU" | Write-Host
    
                [System.Environment]::NewLine
    
                Read-Host -Prompt "Press Enter to return to menu"
    
            }
            '4'{ #Using Asset tag as reference, move bulk devices to desired Workspace OU, write output to host
            
                $Data = Get-CSVandHeaders | Where-Object { $_.PSObject.Properties.Value -ne '' } # This returns the CSV Path as the first line before the asset data, and skips empty cells. I plan to clean this in a future release.

                # Filter out the path CSV Path from $Data
                $FilteredData = $Data[1..($Data.Count)]

                [System.Environment]::NewLine

                $OU = Read-Host "Enter Destination OU Path"

                $FilteredData | ForEach-Object {
    
                    gam update cros query:asset_id:$_ ou "$OU" | Write-Host
                }
    
                [System.Environment]::NewLine
    
                Read-Host -Prompt "Press Enter to return to menu"
            }
            '5'{ #Using Asset tag as reference, wipe user accounts on single device, write output to host
                
                $SingleAsset = Read-Host $lang1
    
                gam issuecommand cros query:asset_id:$SingleAsset command wipe_users doit | Write-Host
    
                [System.Environment]::NewLine
    
                Read-Host -Prompt "Press Enter to return to menu"
            }
            '6'{ #Using Asset tag as reference, wipe user accounts on bulk devices, write output to host
    
                $Data = Get-CSVandHeaders | Where-Object { $_.PSObject.Properties.Value -ne '' } # This returns the CSV Path as the first line before the asset data, and skips empty cells. I plan to clean this in a future release.

                # Filter out the path CSV Path from $Data
                $FilteredData = $Data[1..($Data.Count)]

                $FilteredData | ForEach-Object {
    
                    [System.Environment]::NewLine
    
                    gam issuecommand cros query:asset_id:$_ command wipe_users doit | Write-Host
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
            '9'{
                $Data = Get-CSVandHeaders | Where-Object { $_.PSObject.Properties.Value -ne '' } # This returns the CSV Path as the first line before the asset data, and skips empty cells. I plan to clean this in a future release.

                # Filter out the path CSV Path from $Data
                $FilteredData = $Data[1..($Data.Count)]

                [System.Environment]::NewLine

                $FilteredData | ForEach-Object {
                    $sn = @( gam info cros query:asset_id:$_ serialnumber asset_ID ) | ConvertFrom-String -delimiter "serialNumber: "
                    gam update cros cros_sn $sn.p2 action disable
                }
    
                [System.Environment]::NewLine
    
                Read-Host -Prompt "Press Enter to return to menu"
            }
            '10'{
                $Data = Get-CSVandHeaders | Where-Object { $_.PSObject.Properties.Value -ne '' } # This returns the CSV Path as the first line before the asset data, and skips empty cells. I plan to clean this in a future release.

                # Filter out the path CSV Path from $Data
                $FilteredData = $Data[1..($Data.Count)]

                [System.Environment]::NewLine

                $FilteredData | ForEach-Object {
                    $sn = @( gam info cros query:asset_id:$_ serialnumber asset_ID ) | ConvertFrom-String -delimiter "serialNumber: "
                    gam update cros cros_sn $sn.p2 action reenable
                }
    
                [System.Environment]::NewLine
    
                Read-Host -Prompt "Press Enter to return to menu"
            }
            '11'{ #Using AssetID as reference, Powerwash single device

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
            Write-Host "9. Add individual student to Google Classroom - Single Class" -ForegroundColor White
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
                '3'{ #Using MessageID as reference, delete email from all inboxes

                    [System.Environment]::NewLine
                    $reporter = Read-host "Input email address of reporter"
                    $MessageID = Read-Host "Input Message ID"
                    [System.Environment]::NewLine
                    Write-host "Checking all mailboxes for specified Message ID, and deleting. Operation will take approximately 3 minutes"

                        gam all users delete messages query rfc822msgid:$MessageID doit *> "$env:UserProfile\Documents\'$reporter'_deleteEmail_$(get-date -Format yyyy-MM-dd_HHmm).txt"

                    Get-ChildItem -path $env:UserProfile\Documents -filter "'$reporter'_deleteEmail*.txt" -recurse | Sort-Object CreationTime -Descending | Select-Object -First 1 | select-string -pattern 'Got 1 Messages'
                    [System.Environment]::NewLine
                    Read-Host -Prompt "Press Enter to return to menu"

                }
                '4'{ #Using Sender Address as reference, delete email from all inboxes

                    [System.Environment]::NewLine
                    $reporter = Read-host "Input email address of reporter"
                    $msgsender = Read-Host "Input email address of sender"
                    [System.Environment]::NewLine
                    Write-host "Checking all mailboxes for specified sender address and deleting. Operation will take approximately 3 minutes"

                        gam all users delete messages query "from:$msgsender" doit *> "$env:UserProfile\Documents\'$reporter'_deleteEmail_$(get-date -Format yyyy-MM-dd_HHmm).txt"

                    Get-ChildItem -path $env:UserProfile\Documents -filter "'$reporter'_deleteEmail*.txt" -recurse | Sort-Object CreationTime -Descending | Select-Object -First 1 | select-string -pattern 'Got 1 Messages'
                    [System.Environment]::NewLine
                    Read-Host -Prompt "Press Enter to return to menu"

                }
                '5'{ #Reset password via Google
                    $Username = Read-Host "Input Email Address"
                    $Credential = Read-Host "Input New Password" -AsSecureString

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
                '7'{ #Update Google Classroom owner - bulk
                    
                    $Teacher = read-host "Input Email Address"
                    $Data = Get-CSVandHeaders | Where-Object { $_.PSObject.Properties.Value -ne '' } # This returns the CSV Path as the first line before the asset data, and skips empty cells. I plan to clean this in a future release.

                    # Filter out the path CSV Path from $Data
                    $FilteredData = $Data[1..($Data.Count)]
                    $FilteredData | ForEach-Object {

                        gam course $_ add teacher $Teacher | Write-Host
                        gam update course $_ owner $Teacher | Write-Host

                    }

                    [System.Environment]::NewLine

                    Read-Host -Prompt "Press Enter to return to menu"

                }
                '8'{ #Update Google Classroom owner - single
                    $Teacher = read-host "Input Email Address"
                    $CourseID = read-host "Input Course ID"

                        gam course $CourseID add teacher $Teacher | Write-Host
                        gam update course $CourseID owner $Teacher | Write-Host

                    [System.Environment]::NewLine

                    Read-Host -Prompt "Press Enter to return to menu"
                }
                '9'{ #Add student to Google Classroom
                    $Student = read-host "Input Email Address"
                    $CourseID = read-host "Input Course ID"

                        gam course $CourseID add student $Student | Write-Host

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
