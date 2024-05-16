#####
## To enable scrips, Run powershell 'as admin' then type
## Set-ExecutionPolicy Unrestricted
#####

#################### Transcript Open
$Transcript = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $Transcript | Out-Null
#################### Transcript Open

### Main function header - Put ITAutomator.psm1 in same folder as script
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
$scriptCSV      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".csv"  ### replace .ps1 with .csv
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
$psm1="$($scriptDir)\ITAutomator.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
$psm1="$($scriptDir)\ITAutomator M365.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
############
if (!(Test-Path $scriptCSV))
{
    ######### Template
    "Owner_ID,Delegate_ID,NotifyDelegate,Contacts_Edit,Private_Items,RemoveAllAccess" | Add-Content $scriptCSV
    "owner@contoso.com,delegate@contoso.com,TRUE,FALSE,FALSE,FALSE" | Add-Content $scriptCSV
    ######### Template
	$ErrOut=201; Write-Host "Err $ErrOut : Couldn't find '$(Split-Path $scriptCSV -leaf)'. Template CSV created. Edit CSV and run again.";Pause; Exit($ErrOut)
}
## ----------Fill $entries with contents of file or something
$entries=@(import-csv $scriptCSV)
#$entries
$entriescount = $entries.count
##
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Bulk actions in M365"
Write-Host ""
Write-Host "Use this to add and remove delegates"
Write-Host ""
Write-Host "Owner_ID,Delegate_ID: Owner and Delegate info"
Write-Host "NotifyDelegate      : TRUE (Delegate is sent a summary of access)"
Write-Host "Contacts_Edit       : TRUE (Additionally, contacts are editable)"
Write-Host "Private_Items       : TRUE (Private items too)"
Write-Host "RemoveAllAccess     : TRUE (Delegate is removed)"
Write-Host ""
Write-Host "CSV: $(Split-Path $scriptCSV -leaf) ($($entriescount) entries)"
$entries | Format-Table
Write-Host "-----------------------------------------------------------------------------"
PressEnterToContinue
$no_errors = $true
$error_txt = ""
$results = @()

## ----------Connect/Save Password
$domain=@($entries[0].psobject.Properties.name)[0] # property name of first column in csv
$domain=$entries[0].$domain # contents of property
$domain=$domain.Split("@")[1]   # domain part
#Write-Host "Connect-ExchangeOnline $($domain) ..."
$connected_ok = ConnectExchangeOnline -domain $domain
if (-not ($connected_ok))
{ # connect failed
    Write-Host "[Not connected]"
} # connect failed
else
{ # connected OK
    $processed=0
    $message="$entriescount Entries. Continue?"
    $choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes","&No")
    [int]$defaultChoice = 0
    $choiceRTN = $host.ui.PromptForChoice($caption,$message, $choices,$defaultChoice)
    if ($choiceRTN -eq 1)
    { "Aborting" }
    else 
    { ## continue choices
    $choiceLoop=0
    $i=0        
    foreach ($x in $entries)
    {
        $i++
        write-host "-----" $i of $entriescount $x
        if ($choiceLoop -ne 1)
            {
            $message="Process entry "+$i+"?"
            $choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes","Yes to &All","&No","No and E&xit")
            [int]$defaultChoice = 1
            $choiceLoop = $host.ui.PromptForChoice($caption,$message, $choices,$defaultChoice)
            }
        if (($choiceLoop -eq 0) -or ($choiceLoop -eq 1))
            { # choiceloop
            $processed++
		    ####### Start code for object $x
		    #Owner_Email	Delegate_Email	Contacts_Edit
		    #ajnpierre@roundtableip.com	achu@roundtableip.com	TRUE
		    #######
            $mbfrom=Get-Mailbox $x.Owner_ID
            $mbto=Get-Mailbox $x.Delegate_ID
            ##########
	        $ident = $x.Owner_ID
	        Try {$recip = Get-Recipient -identity $ident}
	        Catch {Write-Warning "$($ident) is not a known email"}
            $MemberEmail=$recip.PrimarySmtpAddress
	        ##########
            $contactsedit = [System.Convert]::ToBoolean($x.Contacts_Edit)
            $notesedit = [System.Convert]::ToBoolean($x.Notes_Edit)
            $privatetoo = [System.Convert]::ToBoolean($x.Private_Items)
            $NotifyDelegate = [System.Convert]::ToBoolean($x.NotifyDelegate)
            $RemoveAccess = [System.Convert]::ToBoolean($x.RemoveAllAccess)
            $DelgateReceivesMeetingNotices = [System.Convert]::ToBoolean($x.DelgateReceivesMeetingNotices)
            ####### Display 'before' info
            Write-host "From: $($mbfrom) [Whether to deliver meeting requests at all (an OWNER setting across all delegates) is not adjusted by this code]"
			Write-host "  To: $($mbto)   "
            #######
            if (-not($mbfrom)) {Write-Warning "From $($mbfrom) is not found, aborting.";Pause;Exit}
            if (-not($mbto)) {Write-Warning "To $($mbto) is not found, aborting.";Pause;Exit}
            ###
            if ($RemoveAccess)
            {
                Write-Host "RemoveAllAccess: True "
                Write-Host "WARNING: Removing all access from $($mbto)" -ForegroundColor Yellow
            }
            # collect flags
            $sharingpermflags = @()
            if ($DelgateReceivesMeetingNotices) {$sharingpermflags+="Delegate"}
            if ($privatetoo)                    {$sharingpermflags+="CanViewPrivateItems"}
            if ($sharingpermflags.count -eq 0) {$strsharingpermflags = "None"} else {$strsharingpermflags = $sharingpermflags -join ","}
            # collect flags
            ### Check for full mailbox permission
            $getperm= Get-MailboxPermission ($mbfrom.PrimarySmtpAddress) -User ($mbto.PrimarySmtpAddress) -ErrorAction silentlycontinue
            if ($getperm)
            {
                Write-host "[Full Mailbox Permission Before]"
                write-host ($getperm |Format-Table Identity,User,AccessRights | Out-String)
                if ($RemoveAccess)
                { #perms exist
                    Remove-MailboxPermission ($mbfrom.PrimarySmtpAddress) -User ($mbto.PrimarySmtpAddress) -AccessRights FullAccess -InheritanceType All -Confirm:$false
                    Write-host "OK: Full Mailbox Permssion removed"
                    $made_changes = $true
                } #perms exist
                else 
                {
                    Write-Host "ERROR: Delegate adjustments aren't applicable whilst full mailbox permission exists. Try removing access first."
                    Start-Sleep -Seconds 3
                    exit
                }
            }
            ### Check for full mailbox permission
            $folders = @("Calendar","Tasks","Contacts","Notes")
            $perms = "Editor"
            foreach ($folder in $folders)
            { #each folder
                Write-Host "Checking: $($folder) of $($mbfrom)"
                # do we deal with this folder?
                $deal_with=$true # default to true for calendar and tasks
                if ($folder -eq "Contacts")
                {
                    $deal_with = $contactsedit -or $RemoveAccess
                }
                if ($folder -eq "Notes")
                {
                    $deal_with = $notesedit -or $RemoveAccess
                }
                if ($folder -eq "Calendar")
                {
                    $deal_with = $true
                }
                #
                if ($deal_with)
                { # deal_with
                    Write-host ("[Permission Before] $($folder) [To $($mbto)]")
                    ## see if user is listed
                    $getperm= Get-MailboxFolderPermission ($mbfrom.PrimarySmtpAddress +":\$folder") -User ($mbto.PrimarySmtpAddress) -ErrorAction silentlycontinue
                    if ($getperm)
                    { #modify existing, Use Set-MailboxFolderPermission
                        Write-Host ($getperm |Format-Table Foldername,User,AccessRights,SharingPermissionFlags | Out-String)
                        $made_changes=$false
                        $perms_ok = ($getperm.AccessRights[0] -eq $perms) 
                        $sharing_ok = $false
                        $remove_perms =  $false
                        if ($RemoveAccess)
                            {$remove_perms = $true}
                        else { # not removeaccess
                            if ($folder -eq "Calendar")
                            { #calendar
                                $sharing_ok = ((Coalesce ($getperm.SharingPermissionFlags,"None")) -eq $strsharingpermflags)
                            } #calendar
                            else
                            { #not calendar 
                                $sharing_ok = $true
                            }
                            if ($perms_ok -and $sharing_ok)
                            {$remove_perms = $false}
                            else
                            {$remove_perms = $true}
                        } # not removeaccess
                        if ($remove_perms)
                        { #perms need changing
                            $made_changes=$true
                             # removing access
                                if ($folder -eq "Calendar")
                                {
                                    Remove-MailboxFolderPermission ($mbfrom.PrimarySmtpAddress +":\$folder") -User ($mbto.PrimarySmtpAddress) -SendNotificationToUser:$NotifyDelegate -Confirm:$false | Out-Null
                                }
                                else {
                                    Remove-MailboxFolderPermission ($mbfrom.PrimarySmtpAddress +":\$folder") -User ($mbto.PrimarySmtpAddress) -Confirm:$false | Out-Null
                                }
                                Write-host ("$($folder): Permission removed")
                        } #perms need changing
                        else
                        { #perms_ok
                            Write-host ("$($folder): already OK")
                            Continue #skip to next folder
                        } #perms_ok
                    }#modify existing
                    if ($RemoveAccess)
                    {
                        Write-host ("$($folder): No prior permission [Already OK]")
                    }
                    else
                    { # Add Access
                        if (!$remove_perms) {Write-host ("$($folder): No prior permission")}
                        $made_changes=$true
                        if ($folder -eq "Calendar")
                        { #calendar
                            Add-MailboxFolderPermission ($mbfrom.PrimarySmtpAddress +":\$folder") -User ($mbto.PrimarySmtpAddress) -AccessRights $perms -SharingPermissionFlags $strsharingpermflags -SendNotificationToUser $NotifyDelegate | Out-Null
                        } #calendar
                        else
                        { #not calendar
                            Add-MailboxFolderPermission ($mbfrom.PrimarySmtpAddress +":\$folder") -User ($mbto.PrimarySmtpAddress) -AccessRights $perms 
                        } #not calendar
                    } # Add Access
                    ##
                    if ($made_changes)
                    { #made changes
                        #### Show rights
                        Write-host ("[Permission After] $($folder) [For All]")
                        Start-Sleep 5 # may need to wait a sec before new permissions show up
                        $getperm_all=Get-MailboxFolderPermission ($mbfrom.PrimarySmtpAddress +":\$folder") -ErrorAction silentlycontinue
                        $getperm_all|Format-Table Foldername,User,AccessRights,SharingPermissionFlags | out-host
                        ####
                    } #made changes
                } # deal_with
                else
                { # don't deal_with
                    Write-host "           [skipping]"
                } # don't deal_with
            } #each folder

            ##### Check send on behalf:START
            $made_changes=$false
            $mbfrom=Get-Mailbox $x.Owner_ID
            # show delegates
            $dlgts = foreach ($delegate in $mbfrom.GrantSendOnBehalfTo) {"[$($delegate)] "}
            if ($null -eq $dlgts) {$dlgts="(none)"}
            Write-host ("$mbfrom [After] Delegates: $($dlgts)")
            # show delegates
            $mbfrom=Get-Mailbox $x.Owner_ID
            if (
                ($mbfrom.GrantSendOnBehalfTo.Contains($mbto.DisplayName)) `
                -or ($mbfrom.GrantSendOnBehalfTo.Contains($mbto.Identity))`
                -or ($mbfrom.GrantSendOnBehalfTo.Contains($mbto.Alias))
                )
            { # delegate exists
                if ($RemoveAccess)
                { # removing access
                    ##### Adjust delegates
                    $newdels = $mbfrom.GrantSendOnBehalfTo
                    $newdels.Remove($mbto.DisplayName) | Out-Null
                    $newdels.Remove($mbto.Identity) | Out-Null
                    $newdels.Remove($mbto.Alias) | Out-Null
                    Get-Mailbox $mbfrom.Identity | Set-Mailbox -GrantSendOnBehalfTo $newdels
                    $made_changes = $true
                } # removing access
                else
                { # adding access
                    Write-Host "Delegates: already OK"
                } # adding access
            } # delegate exists
            else
            { # delegate does not exist
                if ($RemoveAccess)
                { # removing access
                    Write-Host "Delegates: already OK"
                } # removing access
                else
                { # adding access
                    $newdels = $mbfrom.GrantSendOnBehalfTo
                    $newdels.Add($mbto.Identity) | Out-Null
                    Set-Mailbox $mbfrom.PrimarySmtpAddress -GrantSendOnBehalfTo $newdels
                    $made_changes = $true
                } # adding access
            } # delegate does not exist
            if ($made_changes)
            { #made changes
                #### Show delegates
                $mbfrom=Get-Mailbox $x.Owner_ID
                # show delegates
                $dlgts = foreach ($delegate in $mbfrom.GrantSendOnBehalfTo) {"[$($delegate)] "}
                if ($null -eq $dlgts) {$dlgts="(none)"}
                Write-host ("$mbfrom [After] Delegates: $($dlgts)")
                # show delegates
            } #made changes
            ##### Check send on behalf:END

            ####### End code for object $x
            } # choiceloop
        if ($choiceLoop -eq 2)
            {
            write-host ("Entry "+$i+" skipped.")
            }
        if ($choiceLoop -eq 3)
            {
            write-host "Aborting."
            break
            }
        }
    } ## continue choices
    WriteText "Removing any open sessions..."
    Get-PSSession 
    Get-PSSession | Remove-PSSession
    WriteText "------------------------------------------------------------------------------------"
    $message ="Done. " +$processed+" of "+$entriescount+" entries processed. Press [Enter] to exit."
    WriteText $message
    WriteText "------------------------------------------------------------------------------------"
	#################### Transcript Save
    Stop-Transcript | Out-Null
    $date = get-date -format "yyyy-MM-dd_HH-mm-ss"
    New-Item -Path (Join-Path (Split-Path $scriptFullname -Parent) ("\Logs")) -ItemType Directory -Force | Out-Null #Make Logs folder
    $TranscriptTarget = Join-Path (Split-Path $scriptFullname -Parent) ("Logs\"+[System.IO.Path]::GetFileNameWithoutExtension($scriptFullname)+"_"+$date+"_log.txt")
    If (Test-Path $TranscriptTarget) {Remove-Item $TranscriptTarget -Force}
    Move-Item $Transcript $TranscriptTarget -Force
    #################### Transcript Save
} # connected OK
PressEnterToContinue