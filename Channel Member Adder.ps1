##### <INFO> #####

# File:                           Channel Member Adder
# Version:                        0.1.0 Alpha
# Date:                           1 September 2022
# Author:                         Josha
# Website:                        https://github.com/JoshaNL

# Project name:                   Microsoft 365 Scripts
# Repo URL:                       https://github.com/JoshaNL/M365-Scripts


# Copyright© 2022
# All rights reserved.

# LICENSE:
# This file may not be sold, published, and/or distributed in any way, shape, and/or form without prior written agreement of the aforementioned author.

# TERMS:
# No warranties, express or implied, are made with regard to the quality, suitability, or accuracy of this file.
# The use, modification, and/or alteration of this file is wholly and solely at your own risk.
# You acknowledge and accept the risk that the use of this file/script may have unintended consequences.
# You understand that this script is just a quick draft and nowhere near finished.

##### </INFO> #####

###### - PARAMETERS - ######
param(
[string]$Path = ".\list.csv"
)

######## - CONFIG - ########
$VerboseLog = 1
$ErrorLog = 1

####### - VARIABLES - ######
$VerboseLogs = @()
$FailedUsers = @()
$global:LogPath = '.\Script.log'
$global:Logs = $null
$global:Credentials = $null
$global:MFA = $null
$global:CsvCheck = $null
$global:CSV = $null
$global:Team = $null
$global:TeamChannel = $null
$global:CurUser = $null

####### - FUNCTIONS - ######
function LogVerbose($message) {
    if($VerboseLog) {
        $Date = Get-Date
        $LogEntry = "[$date] [INFO] $message"
        Write-Host $LogEntry
        $LogEntry = "$LogEntry`r`n"
        $global:Logs+=$LogEntry
    }
}

function LogError($message) {
    if($ErrorLog) {
        $Date = Get-Date
        $LogEntry = "[$date] [ERROR] $message"
        Write-Host $LogEntry
        $LogEntry = "$LogEntry`r`n"
        $global:Logs+=$LogEntry
    }
}

function SetupEnvironment {
    # Install Modules
    InstallModule("AzureAD")
    InstallModule("MicrosoftTeams")

    # Check Modules
    $AzureADCheck = CheckModule("AzureAD")
    $MicrosoftTeamsCheck = CheckModule("MicrosoftTeams")
    if(-Not $AzureADCheck -Or -Not $MicrosoftTeamsCheck) {
        LogError("AzureAD or MicrosoftTeams modules missing!")
        exit
    }
}

function CheckModule($ModuleName) {
    if(Get-Module -ListAvailable -Name $ModuleName) {
        return 1
    } else {
        return 0
    }
}

function InstallModule($ModuleName) {
    $ModuleCheck = CheckModule($ModuleName)
    if ($ModuleCheck -eq 0) {
        Install-Module -Name $ModuleName
        LogVerbose("Installed $ModuleName module.")
    } else {
        LogVerbose("Found module $ModuleName")
    }
}

function SelectCsv {
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ Filter = 'Comma-separated values (*.csv)|*.csv'}
    $null = $FileBrowser.ShowDialog()
    if($FileBrowser.CheckFileExists -eq 1) {
        $global:CSV = $FileBrowser.FileName.ToString()
        $global:CsvCheck = 1
    } else {
        LogError("Selected CSV file does not exist. Exiting...")
        exit
    }
}

function CheckCsv($Path) {
    if(Resolve-Path -Path $Path) {
        $Path = Resolve-Path -Path $Path
        if(Import-Csv $Path) {
            LogVerbose("Paramater CSV file found, exists and can be imported: $Path")
            $global:CSV = $Path
            $global:CsvCheck = 1
        } else {
            LogError("Parameter CSV file found and exists but can't be imported.")
        }
    } else {
        LogError("Parameter CSV file doesn't exist.")
    }
}

function LoadCsv($Path) {
    $ImportedCsv = $null
    $ImportedCsv = Import-Csv($Path)

    if($ImportedCsv -ne $null -or $ImportedCsv ) {
        LogVerbose("CSV file with path '$Path' successfully imported.")
    }

    $headers = ($ImportedCsv | Get-Member -MemberType NoteProperty).Name
    if($headers -notcontains 'Email'){
        LogError("CSV doesn't have column Email in its header! Exiting...")
        exit
    }

    $ImportedCsv | ForEach-Object {
        if($_.Email.Length -lt 6 -or $_.Email.Contains(' ') -or -not $_.Email.Contains('@') -or -not $_.Email.Contains('.')) {
            LogError("CSV has invalid email addresses! Exiting...")
            exit
        }
    }

    return $ImportedCsv
}

function ConnectAzureAD() {
    $CredsXML = '.\creds.xml'
    if(Resolve-Path -Path $CredsXML -ErrorAction SilentlyContinue) {
        LogVerbose('Found creds.xml in the specified working directory!')
        $global:Credentials = Import-CliXml -Path $CredsXML
        if($global:CurUser = Connect-MicrosoftTeams -Credential $global:Credentials) {
            LogVerbose('Using the credentials in creds.xml to login!')
            return
        } else {
            LogError('An error occurred while connecting to Azure AD using saved creds.')
            LogVerbose('Removing the credentials XML file...')
            Remove-Item $CredsXML
            LogVerbose('Exiting...')
            exit
        }
    }

    while($global:MFA -ne "y" -and $global:MFA -ne "n") {
        if($global:MFA -ne $null) {
            LogError("Wrong value provided ('$global:MFA')! Provide 'y' or 'n'.")
        }
        $global:MFA = Read-Host -Prompt "Does your organization use MFA? (y or n)"
    }
    
    if($global:MFA -eq "n") {
        # Couldn't login using saved credentials
        LogVerbose('Asking for setting new non-MFA credentials...')
        $global:Credentials = Get-Credential
    }
    
    # If credentials set...
    if($global:Credentials -ne $null) {
        LogVerbose('Credentials set! Exporting to creds.xml file...')
        $global:Credentials | Export-CliXml -Path '.\creds.xml'

        # Login using saved credentials
        LogVerbose('Trying to login using the credentials that were set...')
        if($global:CurUser = Connect-MicrosoftTeams -Credential $global:Credentials) {
            LogVerbose('Logged in to AzureAD using the credentials that were set!')
            return
        } else {
            LogVerbose('An error occurred while connecting to AzureAD with saved credentials.')
            LogVerbose('Removing saved credentials...')
            Remove-Item $CredsXML
        }
    }

    # If no credentials set...
    if($global:CurUser = Connect-MicrosoftTeams) {
        LogVerbose('Connected to AzureAD using MFA-supported online AzureAD login screen!')
        return
    }

    # If not able to login...
    LogError('Could not login using saved credentials or MFA-supported online AzureAD login screen!')
    exit
}

function GetTeam() {
    $Team = $null
    $global:CurUserUPN = $global:CurUser.Account.Id.ToString()
    LogVerbose("Opening list of Teams to choose from...")
    $global:Team = Get-Team -User $global:CurUserUPN | Out-GridView -PassThru -Title "Choose team"
    if($global:Team -eq $null) {
        LogError("No team selected. Exiting...")
        exit
    }
    if($Team.ObjectId -ne $null -and $Team.ObjectId -ne "") {
        $global:Team = $Team
        $TeamID = $Team.GroupId
        $TeamName = $Team.DisplayName
        LogVerbose("Team with ID '$TeamID' and name '$TeamName' selected!")
    }
}

function GetChannel() {
    $Channel = $null
    LogVerbose("Opening list of channels to choose from...")
    $Channel = Get-TeamChannel -GroupId $global:Team.GroupId | Out-GridView -PassThru -Title "Choose channel"
    if($Channel -eq $null) {
        LogError("No channel selected. Exiting...")
        exit
    }
    if($Channel.ObjectId -ne $null -and $Channel.ObjectId -ne "") {
        $global:Channel = $Channel
        $ChannelID = $Channel.GroupId
        $ChannelName = $Channel.DisplayName
        LogVerbose("Channel with ID '$ChannelID' and name '$ChannelName' selected!")
    }
}

function EndOfFile() {
    # Open existing log file or create new
    if(Resolve-Path -Path $global:LogPath -ErrorAction SilentlyContinue) {
        $Path = Resolve-Path -Path $global:LogPath
    } else {
        New-Item $global:LogPath > $LogFile
    }

    # Add logs
    Add-Content -Path $global:LogPath -Value $global:Logs
}

########## - EXECUTION - ##########
SetupEnvironment
CheckCsv($Path)
if($global:CsvCheck -ne 1) {
    SelectCsv
}
$LoadedCsv = LoadCsv($global:CSV)
ConnectAzureAD
GetTeam
GetChannel
$TeamUsers = Get-TeamUser -GroupId $global:Team.GroupId
$TeamUsersUPN = $TeamUsers.User

# --- For each entry in the CSV list
$LoadedCsv | ForEach-Object {
    $Failed = 0
    $ListItem = $_
    $ListUser = $ListItem.'Email'.ToString()

    # Check for UPN
    if($TeamUsersUPN -notcontains $ListItem.Email) {
        $Failed = 1
        $FailedUser = $ListUser
        $FailedUsers+=$FailedUser
        LogError("An error occurred while checking UPN: $FailedUser")
    }

    # If UPN found
    if($Failed -eq 0) {
        $Channel = $global:Channel
        $ChannelName = $Channel.DisplayName
        $ChannelID = $Channel.ObjectId
        $ChannelMemberships = Get-TeamChannelUser -GroupId $global:Team.GroupId.ToString() -DisplayName $global:Channel.DisplayName

        LogVerbose("User with UPN '$ListUser' found!")

        ## Group adder
        $GroupAddFail = 1
        if( -not $ChannelMemberships ) {
            $GroupAddFail = 0
        } elseif($ChannelMemberships.User.Contains($ListUser)) {
            LogVerbose("User already in the selected channel (so not added again): $ListUser")
            $GroupAddFail = 1
        } else {
            LogVerbose("User '$ListUser' has channel memberships, but not channel '$ChannelName'.")
            $GroupAddFail = 0
        }
        if($GroupAddFail -eq 0) {
            Add-TeamChannelUser -GroupId $global:Team.GroupId -DisplayName $ChannelName -User $ListUser
            LogVerbose("Tried adding user '$ListUser' to the channel '$ChannelName'.")
        }
    }
} # --- End of ForEach

EndOfFile
exit
