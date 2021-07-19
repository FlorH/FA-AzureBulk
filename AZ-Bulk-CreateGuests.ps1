
<#
.SYNOPSIS
	Creates guest users from Home Warrenty Tenant on firstam tenant in Azure
	
.DESCRIPTION
    This script uses Service Principles and certificates to access each tenant
    It pulls a group from FAHW and looks to see if the user has already been added
    as a Guest to the FA tenant.  If the user already exists, the EmployeeID is compared
    to ensure the email account has not been reused.  Comparing the employeeID will allow for
    updates if a user changes their name.  Once created on the FA side, the user is then added to a group
    on the FA side, which is used to compare and find stale uses to remove.
			
.INPUTS
    Group in FAHW which will control what Guests are created in FA
    Group in FA which contains all FAHW guest which will compare to FAHW to find stale users
    Secret key for both tenants
    Service principal with necessary permissions in FA Tenant to invite, update, and delete Guest users    
	
.OUTPUTS
	Creates an log file---tbd

.LINK
	Brought to you by the EIM Team, another @YesJustFlor production
	
.EXAMPLE
	To be determined may be setup in a runspace in Azure

############################################################################################
ScriptName:		    AZ-Create-Guests
Author:			    Flor H
Date Created: 		April 2019
Reveiwer:			
Date Released		06/10/2021
Current Version:	01.40 
Usage : 			run the ps1 file
***************************************************************************************
Version History
Date		User        Ver	    Description
04/2019 	Flor H	    0.5 	Script Creation in progress 
10/08/2019  Flor H      0.6     Added code to allow for user name change where user kept old name as their email
                                casuing a mismatch in name between tenant UPN and corp mail
    "extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13": "flherrera@firstam.com",
    "extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute12": "4763f9a3-bad0-4118-8fde-94a2575e045a",
    "extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute6": "CORP\\flherrera",
11/15/2019  Flor H.     1.0     Added EagleID for Rep Title
05/2020     Flor H      1.1     Added FCT
01/28/2021  Flor H.     1.2     Added logic for Prometic email suffix for FCT
02/5/2021   Flor H.     1.21    Fixed HW TE & SY Userss who were not getting EA13 set correctly
05/27/2021  Flor H.     1.3     Added Eagle ID for HW, and allowed FCT users without employee IDs to be added
6/10/2021   Flor H.     1.4     Added send mail on critical failure
06/30/2021  Flor H.     1.5     Moved auth to function so they can be called if a token times out
7/13/2021   Flor H.     1.6     When employeeID is missing code is unable to process user, added additional check for Null
                                employeeID for FAHW and Rep Title, they are expected to have employeeID populated
#############################################################################################>
$Error.clear()
#change the drive based on where the script is running
$Script:Drive = "C:"
$Script:LogPath="$Script:Drive\PSLogs\Guests"
if (!(Test-Path "$Script:Drive\PSLogs" )){$null = New-Item -path "$Script:Drive\" -name "PSLogs" -type directory}
if (!(Test-Path "$Script:Drive\PSLogs\Guests" )){$null = New-Item -path "$Script:Drive\PSLogs" -name "Guests" -type directory}

#These URLs are used to access Graph and login, they do not change
$loginURL   = "https://login.microsoftonline.com"
$resource   = "https://graph.microsoft.com"
$GuesInviteUrl="https://myapps.microsoft.com"
$GraphUsersURL      = "https://graph.microsoft.com/v1.0/Users"
$GraphBetaUsersURL  = "https://graph.microsoft.com/Beta/Users/"
$GraphInviteURL   = "https://graph.microsoft.com/v1.0/Invitations/"
$GraphGrpURL = "https://graph.microsoft.com/v1.0/groups/"
$Props       = '?$select=AccountEnabled,displayName,givenName,surname,employeeId,mail,ID,companyName,onPremisesSamAccountName,onPremisesExtensionAttributes'  #properties we are pulling

#
# Change this flag to $true to sync users based on email, do not verify if employeeID matches or exists
#$Script:SyncOnly=$false 
# Change this flag to $True to run the script in verify mode, will not change any data
$Script:TestOnly=$false
# When set to True script will only create new users, no updates or deletions of existing users
$Script:NewUsersOnly=$false
#

##########################################################################################################
#
#  These are the production settings for Azure
#
#################################################################################################

$FATenant   = "firstam.onmicrosoft.com"                        
$FAappID    = "7b5fc675-995d-4a96-8572-84e1884f5b8b"           #SP-FAGuestUserInvite - this is the application we are using to call the graph API 
$FASecret   = "8_WL8aEoY-TtkgtYUjCkm1fLP0_-MY_3x7"    #this is the secret associated with the app

#################################################################################################
#This group is used to hold guest home warranty users
$FAListofHWGuests_GrpOID  ="c300a939-f264-4dbc-a5a0-c5c6292c110e"
$FAListofHWGuests_GrpName = "AAD-Dyn-Guest-FAHWUsers"
$FAtoHWGuestsGrpGraphURL = $GraphGrpURL+$FAListofHWGuests_GrpOID 
$FAtoHWGuestsMembersURL= $FAtoHWGuestsGrpGraphURL+"/Members"+$props 
$HWTenant      = "fahw.onmicrosoft.com"
#$HWTenantID    = "2501abe7-66d5-4c8f-994a-b8e7038ed7d6"            
$HWappID          = "409b4ea8-b0ce-4342-ac31-c5afc83f37e3"          # FirstAm Tenant Sync Service Princple 
$HWSecret       = "5S_~pAEB5ck.2UozC-~W2g2f4RP-rMP2Y1"              # This is the secret associated with the app

$HWCtrlGrpName     = "Azure-Firstam-Tenant"
$HWCtrlGrpOID     = "617f4b15-6cd7-4083-92a7-fb7368efc6ba"
$HWCtrlGrpGraphURL = $GraphGrpURL+$HWCtrlGrpOID 
$HWCtrlGrpMembersGraphURL = $HWCtrlGrpGraphURL+"/Members"+$Props
#$HWCtrlGrpMembersGraphURL = $HWCtrlGrpGraphURL+"/Members"

#################################################################################################
#This group is used to hold guest republic title users
$FAListofRTGuests_GrpOID  ="9333c7f3-2335-42e4-ade8-bdc41060573d"
$FAListofRTGuests_GrpName = "AAD-Dyn-Guest-RepTitleUsers"
$FAtoRTGuestsGrpGraphURL = $GraphGrpURL+$FAListofRTGuests_GrpOID 
$FAtoRTGuestsMembersURL= $FAtoRTGuestsGrpGraphURL+"/Members"+$props 
$RTTenant      = "republictitle.onmicrosoft.com"
#$RTTenantID    = "92a184a7-2c30-40ac-9a52-669c275ca783"            
$RTappID          = "9851d43e-0c02-44db-91e3-928731a17f85"          # FirstAm Tenant Sync Service Princple 
$RTSecret       = 'YeOf-z59ixz1a0~~4XkOAW3V3Q~NFf-gB.'     # This is the secret associated with the app
$RTCtrlGrpName     = "SEC-FA-AllEmployees"
$RTCtrlGrpOID     = "e4e4f1a3-eb53-49de-811c-4f7459ba556b"
$RTCtrlGrpGraphURL = $GraphGrpURL+$RTCtrlGrpOID 
$RTCtrlGrpMembersGraphURL = $RTCtrlGrpGraphURL+"/Members"+$Props

#################################################################################################
#This groups is used to hold guest FCT users
$FAListofFTCGuests_GrpOID  = "b8c67fcc-e7e3-4d9b-ab67-c408e58bde10"
$FAListofFCTGuests_GrpName = "AAD-Dyn-Guest-FCTUsers"
$FAtoFCTGuestsGrpGraphURL = $GraphGrpURL+$FAListofFTCGuests_GrpOID
$FAtoFCTGuestsMembersURL= $FAtoFCTGuestsGrpGraphURL+"/Members"+$props 
$FCTTenant      = "fctca.onmicrosoft.com"
#$FCTTenantID    = "3e9292a5-723b-4746-b589-8ee7b282921b"            
$FCTAppID          = "f7fb0e8a-347e-42f5-87a5-e7ec8051266e"          # FirstAm Tenant Sync Service Princple 
$FCTSecret       = "YeOf-z59ixz1a0~~4XkOAW3V3Q~NFf-gB."
$FCTCtrlGrpName     = "DLG-AAD-PR-FASYNC-DYNC"
$FCTCtrlGrpOID     = "10a394ed-5546-4021-b63f-a96b957f7f25"
#First Canadian Trust new secret 6/30/2021
$FCTSecret       = 'b0t4_fS2xs19BfSE7BC-7-_B2IjSL_S_tn'     # This is the secret associated with the app
#$FCTCtrlGrpName     = "DLG-GL-Azure-FASYNC"
#$FCTCtrlGrpOID     = "fb3d5171-ab35-498d-9357-e884d81c3ba3"
$FCTCtrlGrpGraphURL = $GraphGrpURL+$FCTCtrlGrpOID 
$FCTCtrlGrpMembersGraphURL = $FCTCtrlGrpGraphURL+"/Members"+$Props


#################################################################################################
## Logging Functions
#################################################################################################

$LogDate = Get-Date -UFormat %m%d%Y
$Script:LogFile=$LogDate+"-FAGuests.txt"
$Script:ErrorFile=$LogDate+"-FAGuests-Errors.txt"
$Script:MissingAttribFAHW = $LogDate+"-FAHW-NewUsers.txt"
$Script:MissingAttribRT = $LogDate+"-RepTitle-NewUsers.txt"
$Script:MissingAttribCan = $LogDate+"-FCT-NewUsers.txt"
$Script:UseCorpEmail=$false 
$Script:EA13=$null 

$Script:CentralLog="\\corp.firstam.com\restricted\Enterprise Identity Management\Scripts\PSLogs\Guests"

##################################################
#  Function update log... tbd
##################################################
# SF 1.0 - Add log entry
function fcn_AddLogEntry {
	param($Entry)
	$DateTime = Get-Date -format "yyyy/MM/dd HH:mm"
	Write-Host "$DateTime  $Entry"
	out-file -FilePath $Script:LogPath\$Script:LogFile -InputObject "$DateTime  $Entry" -Append
    If(Test-Path $Script:CentralLog){out-file -FilePath $Script:CentralLog\$Script:LogFile -InputObject "$DateTime  $Entry" -Append}

}

function fcn_AddErrorLogEntry {
	param($Entry)
	$DateTime = Get-Date -format "yyyy/MM/dd HH:mm"
    out-file -FilePath $Script:LogPath\$Script:ErrorFile -InputObject "$DateTime  $Entry" -Append	
    If(Test-Path $Script:CentralLog){out-file -FilePath $Script:CentralLog\$Script:ErrorFile -InputObject "$DateTime  $Entry" -Append}
    fcn_AddLogEntry $Entry 
}

function fcn_AddTenantUserLogEntry {
	param($Entry)
	$DateTime = Get-Date -format "yyyy/MM/dd HH:mm"
    out-file -FilePath $Script:LogPath\$Script:HomeTenantLog -InputObject "$DateTime  $Entry" -Append
    fcn_AddLogEntry $Entry 
}

##################################################
#  Function to handle invoke-webrequest errors
##################################################
Function fcn_ErrorHandling{
    Param($ErrMsg)

    If($ErrMsg.Exception -like "*remote server returned an error: (404) Not Found*"){}
    ElseIf($ErrMsg.Exception -like "*The remote server returned an error: (401) Unauthorized*"){}
    Else {
        #unknown error
    }

}

Function fcn_SendMessage{

    $where = $env:COMPUTERNAME
    $SentFrom = "prod-sa-eim-rprt3@firstam.com"
    $SendToMail = @("flherrera@firstam.com")
    $Subject =  "Guest Invite Script Error on $Where"
    $Body = "Hello 
            <p> There has been a critical erron on the Guest Invite script running on $Where please review the logs</p>

    " 
    
    Send-MailMessage -From $SentFrom -Subject $subject -To $SendToMail -Body $body -SmtpServer mail.firstam.com -BodyAsHtml -Attachments $Script:LogPath\$Script:ErrorFile
    }

##################################################################################################
##################################################################################################
#
#  Functions to get auth token
#
##################################################################################################
##################################################################################################

Function fcn_AuthFA{
    #Authenticate to FA Destination location again so we don't time out
    fcn_AddLogEntry "... Auth to FA Tenant third time"
    $Error.clear()
    $FAbody       = @{grant_type="client_credentials";resource=$resource;client_id=$FAappID;client_secret=$FASecret}
    Try{$FAOauth      = Invoke-RestMethod -Method POST -Uri $loginURL/$FATenant/oauth2/token?api-version=1.0 -Body $FAbody}
    Catch{
        fcn_AddErrorLogEntry "### Auth for FA failed.. stopping ###"
        fcn_AddErrorLogEntry ("### # "+$Error.Exception+" ")
        fcn_AddErrorLogEntry "### "
        fcn_AddErrorLogEntry "### Failure need updated FA Credentials ###"
        fcn_SendMessage
        Return 
    }
    if ($null -ne $FAOauth.access_token){
        $FAauthToken = @{'Authorization'="$($FAOauth.token_type) $($FAOauth.access_token)"}
    }
}

Function fcn_AuthFAHW{
    $Script:HomeTenantLog = $Script:MissingAttribFAHW

    $HWBody       = @{grant_type="client_credentials";resource=$resource;client_id=$HWappID;client_secret=$HWSecret}
    Try{$HWOauth      = Invoke-RestMethod -Method POST -Uri $loginURL/$HWTenant/oauth2/token?api-version=1.0 -Body $HWBody}
    Catch{
        fcn_AddErrorLogEntry "##################################################"
        fcn_AddErrorLogEntry "### Auth for Home Warranty failed.. stopping ###"
        fcn_AddErrorLogEntry ("### # "+$Error.Exception+" ")
        fcn_AddErrorLogEntry "### "
        fcn_AddErrorLogEntry "### Expired Secret need updated Home Warranty Credentials ###"
        fcn_AddErrorLogEntry "##################################################"
        fcn_SendMessage
        $AuthOK=$false
    }
    if ($null -ne $HWOauth.access_token){
        $HWauthToken = @{'Authorization'="$($HWOauth.token_type) $($HWOauth.access_token)"}
        $AuthOK=$true
    }
}

Function fcn_AuthRT{
    $Script:HomeTenantLog = $Script:MissingAttribRT
    $RTBody       = @{grant_type="client_credentials";resource=$resource;client_id=$RTAppID;client_secret=$RTSecret}
    Try{$RTOauth      = Invoke-RestMethod -Method POST -Uri $loginURL/$RTTenant/oauth2/token?api-version=1.0 -Body $RTBody}
    Catch{
        fcn_AddLogEntry " "
        fcn_AddErrorLogEntry "##############################################################"
        fcn_AddErrorLogEntry "### Auth for Repulic Title failed.. STOPPING"
        fcn_AddErrorLogEntry ("### # "+$Error.Exception+" ")
        fcn_AddErrorLogEntry "### "
        fcn_AddErrorLogEntry "### Expired Secret need updated Republic Title Credentials ###"
        fcn_AddErrorLogEntry "##############################################################"
        fcn_SendMessage
        $AuthOK=$false
    }
    if ($null -ne $RTOauth.access_token){
        $RTAuthToken = @{'Authorization'="$($RTOauth.token_type) $($RTOauth.access_token)"}
        $AuthOK=$true
    }
}

Function fcn_AuthFCT{
    $Script:HomeTenantLog = $Script:MissingAttribCan
    $FCTBody       = @{grant_type="client_credentials";resource=$resource;client_id=$FCTAppID;client_secret=$FCTSecret}
    Try{$FCTOauth      = Invoke-RestMethod -Method POST -Uri $loginURL/$FCTTenant/oauth2/token?api-version=1.0 -Body $FCTBody}
    Catch{
        fcn_AddErrorLogEntry "########################################################"
        fcn_AddErrorLogEntry "### Auth for First Canadian Trust failed.. stopping ###"
        fcn_AddErrorLogEntry ("### # "+$Error.Exception+" ")    
        fcn_AddErrorLogEntry "### "
        fcn_AddErrorLogEntry "### Failure need updated FCT Secret ###"
        fcn_AddErrorLogEntry "### Expired Secret need updated FCT Credentials      ###"
        fcn_AddErrorLogEntry "########################################################"
        fcn_SendMessage
        $AuthOK=$false 
    }
    if ($null -ne $FCTOauth.access_token){
        $FCTAuthToken = @{'Authorization'="$($FCTOauth.token_type) $($FCTOauth.access_token)"}
        $AuthOK=$true 
    }
}


##################################################################################################
#  Function to read the Azure group and return members
##################################################################################################
Function fcn_GetAzureGroup{
    Param($tmpAuthToken, $tmpGrpURL, $tmpGrpMembURL, $tmpCtrlGrpName, $tmpTenant)

    #Check source to make sure expected group is found
    [Array]$tmpGrpMembers=$null
    $Error.clear()
    Try{$Group = (Invoke-WebRequest -UseBasicParsing -Headers $tmpAuthToken -Uri $tmpGrpURL -Method GET)
        #$GrpDetail = ConvertFrom-Json -InputObject $Group.Content
    }
    Catch{
        fcn_AddErrorLogEntry "### Unable to find Group $tmpCtrlGrpName on $tmpTenant.... stopping ###"
            [Hashtable]$Script:results = @{IsValid=$false}
            fcn_SendMessage
        Return}

    #Now get the members of the group we will use to see if they need an account created on the destination tenant
    $Error.clear()
    Try {$GroupMembers = (Invoke-WebRequest -UseBasicParsing -Headers $tmpAuthToken -Uri $tmpGrpMembURL -Method GET)
        $ConvertedList = ConvertFrom-Json -InputObject $GroupMembers.Content 
        [Array]$tmpGrpMembers = $ConvertedList.value
    }
    Catch{
        fcn_ErrorHandling $error 
        fcn_AddErrorLogEntry "### Unable to retrieve group members for $tmpCtrlGrpName on $tmpTenant ... Stopping ###"
        [Hashtable]$Script:results = @{IsValid=$false}
        fcn_SendMessage
        Return
    }

    #If the $odata.nextlink is returned, then we have more data, keep doing a GET until all the data is returned
    $nextURL = $ConvertedList."@odata.nextLink"

    #continue to get until there is no nextlink
    if ($null -ne $nextURL){
        Do{
            $NextResults = Invoke-WebRequest -UseBasicParsing -Headers $tmpAuthToken -Uri $nextURL -Method Get -ErrorAction SilentlyContinue 
            $NextMembersList = ConvertFrom-Json -InputObject $NextResults.Content 
            $tmpGrpMembers += $NextMembersList.value
            $nextURL = $NextMembersList."@odata.nextLink"
        }
        While ($null -ne $nextURL)
    }

    [Hashtable]$Script:results = @{IsValid=$true; GrpMembers=$tmpGrpMembers}

}

##############################################################################
#  Function to remove stale users from Firstam tenant
##############################################################################
Function fcn_RemoveUser{
    Param($tmpAuthToken, $tmpGuest)
    
    Write-host "... this is the function to remove user"

    $RemoveUserURL = $GraphUsersURL+$tmpGuest.ID 
    Write-host $RemoveUserURL
    #Invoke-WebRequest -UseBasicParsing -Headers $tmpAuthToken -Uri $RemoveUserURL -Method Delete
}

##############################################################################
#  Function get a users group membership and check for azure only groups
##############################################################################
Function fcn_GetGrpMembership{
    Param($FALookupDetail)

    $lookupBetaURL = $GraphBetaUsersURL+$FALookupDetail.ID+"/memberof"

    #Try and lookup the guest user in FA first to see if the email already exist
    fcn_AddLogEntry ("... . Get existing FA Group memberships")
    Try{$GrpLookup = Invoke-WebRequest -UseBasicParsing -Headers $FAauthToken -Uri $lookupBetaURL -Method GET
        [Array]$FAGrpDetail = (ConvertFrom-Json -InputObject $GrpLookup.Content).value 
    }
    Catch {
        #throw error.. in testing no error was given, returned no data instead if we get an error something else is going on
        fcn_AddErrorLogEntry "... # Unknown error retrieving group membership, skipping rename"
        [Hashtable]$Script:results = @{IsValid=$false; FAGrps=$null} 
        fcn_SendMessage
        Return}

    fcn_AddLogEntry ("... . User is a member of "+$FAGrpDetail.count+" groups")

    [Hashtable]$Script:results = @{IsValid=$true; FAGrps=$FAGrpDetail}    

}

##############################################################################
#  Function after user is recreated, add the user back to the azure groups
##############################################################################
Function fcn_ApplyGrpMembership{
    Param($tmpGuest)

    Write-host "... add user back to groups"

}

##############################################################################
#  Function Update the Guest extended attributes
##############################################################################
Function fcn_UpdateGuestAttributes{
    Param($NewUserOID, $tmpGuest, $company, $tmpMail, $tmpLookupDetail)

    fcn_AddLogEntry ("... .. Azure User ObjectID: "+$NewUserOID) 

    #IF($null -eq $Script:EagleID){
    #    If($company -eq "Republic Title"){
    #       fcn_AddLogEntry ("... .. Use EgleID from home tenant: "+$tmpGuest.onPremisesExtensionAttributes.extensionAttribute10)
    #        $Script:EagleID  = $tmpGuest.onPremisesExtensionAttributes.extensionAttribute10
    #    }
    #    ElseIf($Company -eq "FAHW"){
    #        fcn_AddLogEntry ("... .. Use EgleID from home tenant: "+$tmpGuest.onPremisesExtensionAttributes.extensionAttribute12)
    #        $Script:EagleID = $tmpGuest.onPremisesExtensionAttributes.extensionAttribute12
    #    }
    #}

    # EA13 is SSO email address
    #$Entry = ("... .. Use as EA13 "+$Script:EA13)
    #fcn_AddLogEntry $Entry 
    #If($Script:NewUser){out-file -FilePath $Script:LogPath\$Script:LogNewUsers -InputObject "$DateTime  $Entry" -Append}

    If($Script:EIDMatch){
        #Found employee ID match, no update required for now
    }
    Else{
        If($null -eq $tmpGuest.EmployeeID){

            If (($company -eq "Republic Title") -and ($tmpMail -notlike "*@reuniontitle.com")){
                out-file -FilePath $Script:LogPath\$Script:HomeTenantLog -InputObject "$DateTime  $Entry" -Append
                fcn_AddLogEntry ("... .. Reunion Title does not have Employee IDs yet")
            }
            ElseIf(($null -eq $Script:UserHomeDetail.EmployeeId) -and ($Tenant -eq $HWTenant) -and (($SAM -like "TE-*") -or ($SAM -like "SY-*"))){
                fcn_AddLogEntry ("... %% Missing EmployeeID on home tenant $company for TE or SY user ok to continue")
            }
            Else{
                $Entry = ("%%% %% Missing EmployeeID on home tenant $company for "+$tmpGuest.GivenName+" "+$tmpGuest.SurName+" "+$tmpMail) 
                fcn_AddErrorLogEntry $Entry 
                fcn_AddTenantUserLogEntry $Entry 
                out-file -FilePath $Script:LogPath\$Script:HomeTenantLog -InputObject "$DateTime  $Entry" -Append
                fcn_AddErrorLogEntry ("%%% %% EmployeeID on FA Tenant is "+$FALookupDetail.EmployeeID)
                fcn_AddErrorLogEntry ("%%% %% guest extended attributes not updated") 
                    Return 
            }
        }
    }

    $UserBetaURL = $GraphBetaUsersURL+$NewUserOID
    $JsonUpdate = @{
        surname=$tmpGuest.SurName;
        GivenName=$tmpGuest.GivenName;
        employeeId=$tmpGuest.EmployeeID;
        extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute12 = $Script:EagleID;
        extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13 = $Script:EA13;
        companyName=$company;        
    } | ConvertTo-Json 
    
    fcn_AddLogEntry ("... . These are the values to be updated")
    fcn_AddLogEntry ("... .. givenname           = "+$tmpGuest.GivenName)
    fcn_AddLogEntry ("... .. surname             = "+$tmpGuest.SurName)
    fcn_AddLogEntry ("... .. employeeId          = "+$tmpGuest.EmployeeID)
    fcn_AddLogEntry ("... .. EagleID             = "+$Script:EagleID)
    fcn_AddLogEntry ("... .. EA13 SSO            = "+$Script:EA13)
    fcn_AddLogEntry ("... .. companyName         = "+$company)

    $Error.clear()
    $null=$FALookupDetail
    Try {
        If($Script:TestOnly){
            fcn_AddLogEntry ("*** . Test only flag is on - skip PATCh to apply updates")
            $FALookupDetail = $tmpLookupDetail
        }
        Else{
            $UpUser = (Invoke-WebRequest -UseBasicParsing -Headers $FAauthToken -Uri $UserBetaURL -Method PATCH -body $JsonUpdate -ContentType "application/json")
        }
    }
    Catch{
        #some error on update, capture and continue
        fcn_AddErrorLogEntry ("%%% Unable to update additional attributes for "+$tmpGuest.GivenName+" "+$tmpGuest.SurName+" manually verify")
        [Hashtable]$Script:results = @{IsValid=$false; RC=$Error}            
        Return 
    }
  
    If($Script:TestOnly){fcn_AddLogEntry ("... . Skip validation")}
    Else{
        If($UpUser.StatusCode -eq "204"){
            fcn_AddLogEntry ("... . ")
            fcn_AddLogEntry ("... . Good status code returned - "+$UpUser.StatusCode+" now get user to validate")
            Try{$ValUser = (Invoke-WebRequest -UseBasicParsing -Headers $FAauthToken -Uri $UserBetaURL -Method GET)
                $ValDetail = ConvertFrom-Json -InputObject $ValUser.Content
                }
            Catch{
                fcn_AddErrorLogEntry ("### # Unable to retrieve updated user detail for verification, manually verify")
                Return 
            }
        }
        Else{
            fcn_AddErrorLogEntry ("### # Unexpected error code"+$UpUser.StatusCode+" unable to validate user, manually verify")
            Return
        }

        If($ValDetail.EmployeeID -ne $tmpGuest.EmployeeID){
            fcn_AddErrorLogEntry ("%%% . EmployeeID "+$ValDetail.EmployeeID+" not set correctly on FA Tenant for "+$tmpMail+" manually verify") 
        }
        Else{fcn_AddLogEntry ("... . EmployeeID has been set  = "+$ValDetail.EmployeeID)}

        #fcn_AddLogEntry ("... . Validated EA13      = "+$ValDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13)
        If($ValDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13 -ne $Script:EA13){
            fcn_AddErrorLogEntry ("%%% . EA13 not set correctly on FA Tenant for "+$tmpMail+" manually verify") 
        }
        Else{fcn_AddLogEntry ("... . EA13 has been set        = "+$Script:EA13)}

        #fcn_AddLogEntry ("... . Validated Company   = "+$ValDetail.companyName)
        If($ValDetail.companyName -ne $company){
            fcn_AddErrorLogEntry ("%%% . Company not set correctly on FA Tenant for "+$tmpMail+" manually verify") 
        }
        Else{fcn_AddLogEntry ("... . Company has been set     = "+$company)}

        If(($company -eq "Republic Title") -or ($Company -eq "FAHW")){
            If($ValDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute12 -ne $Script:EagleID){
                fcn_AddErrorLogEntry ("%%% . Eagle ID not set correctly on FA Tenant for "+$tmpMail+" manually verify") 
            }
            Else{fcn_AddLogEntry ("... . Eagle ID has been set    = "+$Script:EagleID)}
        }

    }
}

##############################################################################
# Function to Add Guest to FA Tenant
##############################################################################
Function fcn_AddGuestToFA{
    Param($tmpGuest, $tmpemail, $company)
    $null = $Content; $NewUserOID="x"; $null=$ValUser; $null=$newUser; $null=$JsonInvite; $null=$JsonUpdate; $IsValid=$true  

    $JsonInvite = @{
        invitedUserEmailAddress=$Script:InviteMail;
        inviteRedirectUrl=$GuesInviteUrl;
        invitedUserDisplayName=$tmpGuest.DisplayName;
        sendInvitationMessage=$false;        
    } | ConvertTo-Json 

    fcn_AddLogEntry ("... . These are the values which will be used to create new Guest User")
    $Entry = ("... . Invited User email  = "+$Script:InviteMail)
    fcn_AddTenantUserLogEntry $Entry 
    $Entry = ("... . Displayname         = "+$tmpGuest.DisplayName)
    fcn_AddTenantUserLogEntry $Entry 

    $Error.clear()
    $Script:NewUser=$false 
    If($Script:TestOnly){
        fcn_AddLogEntry ("*** . Test only flag is on - skip Creating user for now")
        $Script:NewUser=$false
        [Hashtable]$Script:results = @{IsValid=$false; RC="TestFlag"}
    }
    Else{
        $NewUser=$null; [Array]$Content=$null; $NewUserOID=$null 
        fcn_AddTenantUserLogEntry ("... . Creating "+ $tmpGuest.DisplayName +" as Guest on FA Tenant")
            
        Try{$NewUser = (Invoke-WebRequest -UseBasicParsing -Headers $FAauthToken -Uri $GraphInviteURL -Method POST -body $JsonInvite)}
        Catch{
            #some error on create user capture and continue
            fcn_AddErrorLogEntry "### # Error thrown during user creation, manually verify ###" 
            fcn_AddErrorLogEntry("### # "+$Script:InviteMail+" was not created as guest on FA Tenant")
            fcn_AddErrorLogEntry ("### # "+$Error.Exception+" was not created as guest on FA Tenant")
                [Hashtable]$Script:results = @{IsValid=$false; RC=$Error} 
                Return          
        }
    
        $Content = ConvertFrom-Json -InputObject $NewUser.content
        If($Content.count -eq 0){
            fcn_AddLogEntry ("### # User not created correctly")
            [Hashtable]$Script:results = @{IsValid=$false; RC=$Error}       
            Return 
        }
        
        $NewUserOID = $content.invitedUser.id
        fcn_AddLogEntry ("... . Newly created user Object ID: $NewUserOID")
        #fcn_AddLogEntry ("... . Invited eMail address: "+$content.invitedUserEmailAddress)

        $UserURL=$null; $ValUser=$null; [Array]$UserDetail=$null
        $UserURL = $GraphBetaUsersURL+$NewUserOID
        fcn_AddLogEntry ("... . Get newly created user from FA Tenant and display")
        $ValUser = (Invoke-WebRequest -UseBasicParsing -Headers $FAauthToken -Uri $UserURL -Method GET)
        $UserDetail = ConvertFrom-Json -InputObject $ValUser.Content
                        
        If($UserDetail.count -eq 0){
            fcn_AddLogEntry ("### # User not created correctly")
            Write-host $JsonInvite -ForegroundColor Yellow 
            Start-sleep 2
            Return 
            #validate user was created                
        }
        Else{
            fcn_AddLogEntry ("... . Validated email address: "+$UserDetail.mail)
            fcn_AddLogEntry ("... .               User type: "+$UserDetail.userType)
            fcn_AddLogEntry ("... .                     UPN: "+$UserDetail.userPrincipalName)
            $Script:NewUser=$true 
        }
    }
    
    #now update the additional attributes
    fcn_AddLogEntry "... . call function to update guest attributes"
    fcn_UpdateGuestAttributes $NewUserOID $tmpGuest $company $tmpemail $null 
    fcn_AddLogEntry (".... Guest update completed")

}

##############################################################################
# Function Figure out what UPN should be so users can login
##############################################################################
Function fcn_GetUPN{
    Param($tmpUser, $tmpSAM, $tmpAuthToken)
    
    $UserURL = $GraphBetaUsersURL+$tmpuser.id 
    fcn_AddLogEntry ("... . get additional detail from home tenant for $tmpSAM")
    Try{$ValUser = (Invoke-WebRequest -UseBasicParsing -Headers $tmpAuthToken -Uri $UserURL -Method GET)}
    Catch{
        [Hashtable]$Script:results = @{IsValid=$false; newMail=$null}    
        Return
    }
    $UserDetail = ConvertFrom-Json -InputObject $ValUser.Content
    fcn_AddLogEntry ("... . UPN in home tenant is "+$UserDetail.userPrincipalName)
    $tmpNewMail = $UserDetail.userPrincipalName 
    [Hashtable]$Script:results = @{IsValid=$true; newMail=$tmpNewMail}

}
##############################################################################
# Function Set the mail to the tenant UPN and verify address is not in use
##############################################################################
Function fcn_SetMailToUPN{
    Param($tmpUser, $tmpSAM, $tmpMailName, $tmpEmail, $tmpAuthToken)

    $ValUser=$null; $UserDetail=$null 
    $UserURL = $GraphBetaUsersURL+$tmpuser.id 

    fcn_AddLogEntry ("... . get additional detail from home tenant for $tmpSAM")
    $ValUser = (Invoke-WebRequest -UseBasicParsing -Headers $tmpAuthToken -Uri $UserURL -Method GET)
    $UserDetail = ConvertFrom-Json -InputObject $ValUser.Content
            
    fcn_AddLogEntry ("... . UPN from home tenant "+$UserDetail.userPrincipalName+ " will be used for Invite")
    
    #was user created with incorrect invite email?
    $lookupBetaURL = $null; $Lookup=$null; $FALookupDetail=$null 
    $mailfilter = "`$filter=mail eq `'$Script:InviteMail'"
    $lookupBetaURL = $GraphBetaUsersURL+'?'+$mailfilter

    #Try and lookup the guest user in FA to see if they were created with incorrect email address
    fcn_AddLogEntry ("... . Was guest user created with an incorrect email of "+$Script:InviteMail+" created?")
    ####
    #### Need to add logic in the event the email address matching is valid for an FA user
    ####
    Try{$Lookup = Invoke-WebRequest -UseBasicParsing -Headers $FAauthToken -Uri $lookupBetaURL -Method GET
        [Array]$FALookupDetail = (ConvertFrom-Json -InputObject $Lookup.Content).value 
        #Write-host $FALookupDetail -ForegroundColor Yellow        
    }
    Catch {
        #throw error.. in testing no error was given, returned no data instead if we get an error something else is going on
        fcn_AddErrorLogEntry ("... fcn_SetMailtoUPN unknown error in get by email filter for "+$Script:InviteMail)
        fcn_AddTenantUserLogEntry ("... fcn_SetMailtoUPN unknown error in get by email filter for "+$Script:InviteMail)
        $cntSkip++
        Return 
    }
 
    [Hashtable]$Script:results = @{IsValid=$True}
    If($FALookupDetail.count -eq 0){
        fcn_AddLogEntry ("... . Bad Guest email "+$Script:InviteMail+" not found in FA tenant")
        $Script:InviteMail = $UserDetail.userPrincipalName
    }
    Else{
        [Hashtable]$Script:results = @{IsValid=$false}
        #$Entry = ("### # $Script:InviteMail found in FA tenant for "+$FALookupDetail.userType+" and needs to be reviewed ###")
        #fcn_AddErrorLogEntry $Entry 
        #out-file -FilePath $Script:LogPath\$Script:LogNewUsers -InputObject "$DateTime  $Entry" -Append

        If($FALookupDetail.userType -eq "Guest"){
            $Entry = ("### # $Script:InviteMail is assigned to a guest and must be removed ###")
            out-file -FilePath $Script:LogPath\$Script:LogNewUsers -InputObject "$DateTime  " -Append
            out-file -FilePath $Script:LogPath\$Script:LogNewUsers -InputObject "$DateTime  $Entry" -Append
            out-file -FilePath $Script:LogPath\$Script:LogNewUsers -InputObject "$DateTime  " -Append
            fcn_AddErrorLogEntry $Entry 
        }
        Else{
            fcn_AddLogEntry ("... . This is the email address is used by: ")
            fcn_AddLogEntry ("... . Mail: "+$FALookupDetail.mail)
            fcn_AddLogEntry ("... . UPN: "+$FALookupDetail.onPremisesUserPrincipalName)
            fcn_AddLogEntry ("... . EA6: "+$FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute6)
            fcn_AddLogEntry ("... . EA13: "+$FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13)
            fcn_AddLogEntry ("... . Employee ID: "+$FALookupDetail.EmployeeID)
            fcn_AddLogEntry ("... . Manually verify")
           
        }
    }

}
##############################################################################
# HW will use firstam.com suffix in EA13 but needs mail set to UPN
##############################################################################
Function fcn_CheckFAHWeMail{
    Param($Guest, $tmpAuthToken)

    # FAHW Users are created in their tenant with a firstam.com email address

    # can't create guests users with our domain suffix i our tenant, use their tenants 
    # suffix for their email

    [Hashtable]$Script:results = @{IsValid=$True}
    $SAM = $Guest.onPremisesSamAccountName
    $Script:EA13 = $Guest.OnPremisesExtensionAttributes.extensionAttribute13
    $email = $Guest.Mail
    $tmpMailName = $email.split("@")[0]
    fcn_AddLogEntry ("... . Home Warranty user need additional mail validation, FA Guest email is $email")
    
    # Do the SamAccountName and Mail name match
    If($tmpMailName -ne $SAM){        
        #fcn_AddErrorLogEntry ("### "+$Guest.DisplayName)
        $Script:InviteMail = $email.replace("@firstam.com","@fahw.com")
        fcn_AddLogEntry ("... % eMail name mismatch for "+$Guest.DisplayName)
        fcn_AddLogEntry ("... % OnPremiseSamAccountName $SAM and email name "+$email+" do not match")
        fcn_SetMailToUPN $Guest $SAM $tmpMailName $email $tmpAuthToken
            $IsValid = $Script:results.IsValid
        If(!($Isvalid)){
            [Hashtable]$Script:results = @{IsValid=$false}
            fcn_AddErrorLogEntry ("### # skipping user ###")
            Continue 
        }
    }
    ElseIf($email -like "*@fahw.com"){
        #tenant only users have an @fahw email address.
        fcn_AddLogEntry ("... . user already has fahw email: $email")
        
        $Script:InviteMail =  $email
    }
    ElseIf($email -like "*@firstam.com"){
        #Change the email from to fahw.
        fcn_AddLogEntry ("... . user has firstam suffix "+$email+" change to @fahw.com so we can reference Azure Tenant")
        $Script:EA13=$email 
        #fcn_AddLogEntry ("... . Will use "+$Script:EA13+" as Corporate email for EA13 if attribute is not populated")
        $Script:InviteMail = $email.replace("@firstam.com","@fahw.com")
        fcn_AddLogEntry ("... . email changed: "+$Script:InviteMail)
                      
    }
    
}

############################################################################################################################
############################################################################################################################
#  Cycle through the list of users from the outside tenant
############################################################################################################################
############################################################################################################################
Function fcn_ProcessUsers{
    Param($GuestList, $Tenant, $tmpAuthToken)

    If($Tenant -eq $HWTenant){$company="FAHW"}
    ElseIf($Tenant -eq $RTTenant){$Company="Republic Title"}
    ElseIf($Tenant -eq $FCTTenant){$Company="FCT"}
    Else {
        #skip and return error
        fcn_AddLogEntry ("##### Unknown Tenant "+$Tenant+" STOPPING #####")
        $company=$null 
        Return
    }
    
    out-file -FilePath $Script:LogPath\$Script:HomeTenantLog -InputObject "$DateTime  Starting $Company" -Append 

    #CRUD Counts
    $cntC=0; $cntMissMatch=0; $cntU=0; $cntD=0; $cntSkip=0; $cntFound=0; $cntError=0;
    ForEach($Guest in $GuestList){
        #Start-Sleep 5
        WRite-Host ""
        fcn_AddLogEntry ("... ")
        fcn_AddLogEntry ("... Checking "+$Guest.DisplayName+" - "+$Guest.Mail+" in "+$Tenant)
        $cnt++; $Script:InviteMail=$null; $Script:UserHomeDetail = $Null; $Script:Rename=$false
        $email = $Guest.Mail
        
        #user is missing email skip 
        If($null -eq $email){
            $Entry = ("### # User "+$Guest.DisplayName+" missing email attribute "+$Guest.Mail+" ###")
            fcn_AddErrorLogEntry $Entry 
            out-file -FilePath $Script:LogPath\$Script:HomeTenantLog -InputObject "$DateTime  $Entry" -Append
            $Entry = ("### # Unable to verify, user won't be invited as guest ###")
            fcn_AddErrorLogEntry $Entry 
            out-file -FilePath $Script:LogPath\$Script:HomeTenantLog -InputObject "$DateTime  $Entry" -Append
            $cntSkip++; Continue}

        #lookup user on home tenant to get more information
        $UserURL = $GraphBetaUsersURL+$Guest.ID
                    
        #Skip if not found in home tenant or have other issue
        fcn_AddLogEntry ("... .. Get full details from home tenant")
        Try{$Script:tmpUser = (Invoke-WebRequest -UseBasicParsing -Headers $tmpAuthToken -Uri $UserURL -Method GET)
            $script:UserHomeDetail = (ConvertFrom-Json -InputObject $Script:tmpUser.Content)
        }
        Catch{$Script:UserHomeDetail = $Null
            fcn_AddLogEntry ("... . not found in home tenant")
            Continue}

        $Script:EagleID=$null
        $SAM = $Script:UserHomeDetail.onPremisesSamAccountName
        fcn_AddLogEntry ("... . Found "+$Script:UserHomeDetail.DisplayName+" in "+$Tenant)
        fcn_AddLogEntry ("... .. Assigned email in home tenant:  $email")
        fcn_AddLogEntry ("... .. on Premise SamAccountName:      $SAM")
        fcn_AddLogEntry ("... .. Cloud UPN is                    "+$Script:UserHomeDetail.userPrincipalName)
        fcn_AddLogEntry ("... .. Employee ID:                    "+$Script:UserHomeDetail.EmployeeID)
        fcn_AddLogEntry ("... .. Primary email (EA13):           "+$Script:UserHomeDetail.OnPremisesExtensionAttributes.extensionAttribute13)
        If($Tenant -eq $RTTenant){
            fcn_AddLogEntry ("... .. Eagle ID Rep Title:             "+$Script:UserHomeDetail.OnPremisesExtensionAttributes.extensionAttribute10)
            $Script:EagleID = $Script:UserHomeDetail.OnPremisesExtensionAttributes.extensionAttribute10
        }
        ElseIf($Tenant -eq $HWTenant){
            fcn_AddLogEntry ("... .. Eagle ID FAHW:                  "+$Script:UserHomeDetail.OnPremisesExtensionAttributes.extensionAttribute12)
            $Script:EagleID = $Script:UserHomeDetail.OnPremisesExtensionAttributes.extensionAttribute12
        }
        Else{
            fcn_AddLogEntry ("... .. Eagle ID FCT:                    Not in use")
            $Script:EagleID=$null}
        
        fcn_AddLogEntry ("... .. figure out invite email")
                      
        #Figure out what to use as the invite email
        $Script:AdditionalProxy=$false
        If($Tenant -eq $RTTenant){
            If(($email -like "*@republictitle.com") -or ($email -like "*@reuniontitle.com") -or ($email -like "*@hklegal.net")){
                $Script:InviteMail = $email
                $FALookupEmail = $Email
                $Script:EA13 = $email
            }
            Else{
                #expecting users to have first republic email address               
                fcn_AddErrorLogEntry ("%%% . Republic Title - expecting republictitle suffix on email " + $email+" invalid email suffix Skipping user %%%")
                $cntSkip++; Continue
            }
        }
        ElseIf($Tenant -eq $HWTenant){        
            #check email for FAHW
            fcn_CheckFAHWeMail $Script:UserHomeDetail $tmpAuthToken
                $IsValid = $Script:results.IsValid
                If(!($Isvalid)){
                    fcn_AddErrorLogEntry ("### # Remove existing guest account for "+$Guest.DisplayName+" before creating new")
                    Continue 
                }
            $FALookupEmail = $Script:InviteMail
        }
        ElseIf($Tenant -eq $FCTTenant){
            
            If(($email -like "*@fct.ca") -or ($email -like "*@promeric.com")){ 
                If($email -like "*@promeric.com"){$Script:AdditionalProxy=$true}   #add both firstcdn and promeric to additional email field
                $Script:EA13 = $email
                $FALookupEmail = $Script:UserHomeDetail.userPrincipalName
                $Script:InviteMail = $Script:UserHomeDetail.userPrincipalName
                $Company="FCT"
                #$Script:EagleID=$null
            }
            Else{
                fcn_AddErrorLogEntry ("%%% . First Canadian Title - expecting @fct.ca suffix on email " + $email+" invalid email suffix Skipping user %%%")
                Continue
            } 
        }
        Else{
            fcn_AddLogEntry ("... Unknown or unexpected tenant $tenant")
            fcn_AddLogEntry ("... skipping user $email")
            $cntSkip++; Continue
        }
         
        fcn_AddLogEntry ("... .. Invite email will be: "+$Script:InviteMail)

        #Set variables to lookup guest user on FA tenant to see if they already exist or if email is already in use
        $lookupBetaURL = $null; $Lookup=$null; $FALookupDetail=$null 
        $mailfilter = "`$filter=mail eq `'$FALookupEmail'"
        $lookupBetaURL = $GraphBetaUsersURL+'?'+$mailfilter

        #Try and lookup the guest user in FA first to see if the email already exist
        fcn_AddLogEntry ("... . Lookup "+$Guest.DisplayName+" with email of "+$FALookupEmail+ " on FA Tenant")
        Try{$Lookup = Invoke-WebRequest -UseBasicParsing -Headers $FAauthToken -Uri $lookupBetaURL -Method GET
            [Array]$FALookupDetail = (ConvertFrom-Json -InputObject $Lookup.Content).value 
        }
        Catch {
            #throw error.. in testing no error was given, returned no data instead if we get an error something else is going on
            fcn_AddErrorLogEntry ("... Lookup Guest in FA - unknown error in get by email filter with FA eMail"+$FALookupEmail)
            fcn_AddTenantUserLogEntry ("... Lookup Guest in FA - unknown error in get by email filter with FA eMail"+$FALookupEmail)
            $cntSkip++
            Continue
        }
 
        #Did we find a user with the same email?
        If($FALookupDetail.mail -eq $FALookupEmail){
            # we found an email match
            fcn_AddLogEntry ("... . eMail address found at FA")
                
            If($Script:NewUsersOnly){
                fcn_AddLogEntry ("... . Flag has been set to create new users only, skip attribute updates")
                Continue}
                
            fcn_AddLogEntry ("... . Compare home tenant employeeID to FA tenant before making changes")
            fcn_AddLogEntry ("... .. Guest EmployeeID: "+$FALookupDetail.EmployeeID)
            fcn_AddLogEntry ("... .. $company EmployeeID: "+$Script:UserHomeDetail.EmployeeID)
    
            If($FALookupDetail.EmployeeID -ne $Guest.EmployeeID){
                #this user had a different employee ID, must be a different user delete existing and create new
                $Script:EIDMatch = $false
                fcn_AddLogEntry "... . employee IDs don't match, email reassigned, delete existing and create new (PENDING)"
                fcn_AddLogEntry "... . no additional attributes checked"
                $cntU++
    #           fcn_RemoveUser $Guest
    #           $IsValid = $Script:results.IsValid
                
                If($IsValid){
                    #user was deleted, create new
                    #fcn_AddLogEntry ("... . older user was removed, new user created")
                    #fcn_AddGuestToFA $Guest $email $company    
                }
                Else{
                    #user was not deleted, unable to create new, check manually
                    #fcn_AddLogEntry ("... . eMail was reassigned, unable to delete old user and create new, manually verify")
                    $cntError++          
                }
            }       #employee ID does not match
            Else{
                If($null -eq $Guest.EmployeeID){fcn_AddLogEntry "... % Missing Employee IDs from home tenant"
                    $Script:EIDMatch = $false}
                Else{fcn_AddLogEntry "... . employee IDs match, same user check EA12 and EA13"
                    $Script:EIDMatch = $true}
    
                ######################################################################################
                #check to see if we are missing any required attributes EA13=Mail or EA12=Eagleid
                ######################################################################################
                $UpdateUser=$false 
                # EA12 is Eagle ID - right now only Republic Title has EagleID on EA10
                If($company -eq "Republic Title"){
                    If ($FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute12 -eq $Script:UserHomeDetail.onPremisesExtensionAttributes.extensionAttribute10){
                        fcn_AddLogEntry ("... .. EagleID on match FA tenant is "+$FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute12)
                        fcn_AddLogEntry ("... .. EagleID from $company tenant is "+$Script:UserHomeDetail.onPremisesExtensionAttributes.extensionAttribute10)
                    }
                    Else{    
                        fcn_AddLogEntry ("... % EalgeID Missing or incorrect, update required")
                        fcn_AddLogEntry ("... %% EagleID on FA tenant is "+$FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute12)
                        fcn_AddLogEntry ("... %% EagleID from $company tenant is "+$Script:UserHomeDetail.onPremisesExtensionAttributes.extensionAttribute10)
                        $UpdateUser=$true
                    }
        
                }
                #Check to see if EA13 (mail) is correct
                ElseIf($Tenant -eq $HWTenant){
                    If ($FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13 -eq $Script:EA13){
                        fcn_AddLogEntry ("... .. EA13 match FA tenant is "+$FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13)
                        fcn_AddLogEntry ("... .. EA13 from $company tenant is "+$Script:EA13)
                    }
                    Else{
                        fcn_AddLogEntry ("... % EA13 missing or incorrect, update required")
                        fcn_AddLogEntry ("... %% EA13 on FA tenant is "+$FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13)
                        fcn_AddLogEntry ("... %% EA13 from $company tenant is "+$Script:EA13)                        
                        $UpdateUser=$true                        
                    }
                    #Now check Eagle ID EA12
                    If ($FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute12 -eq $Script:UserHomeDetail.onPremisesExtensionAttributes.extensionAttribute12){
                        fcn_AddLogEntry ("... .. EagleID match FA tenant is "+$FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute12)
                        fcn_AddLogEntry ("... .. EagleID from $company tenant is "+$Script:UserHomeDetail.onPremisesExtensionAttributes.extensionAttribute12)
                    }
                    Else{    
                        fcn_AddLogEntry ("... % EalgeID Missing or incorrect, update required")
                        fcn_AddLogEntry ("... %% EagleID on FA tenant is "+$FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute12)
                        fcn_AddLogEntry ("... %% EagleID from $company tenant is "+$Script:UserHomeDetail.onPremisesExtensionAttributes.extensionAttribute12)
                        $UpdateUser=$true
                    }
                }                
                ElseIf($Tenant -eq $FCTTenant){
                    If ($null -eq $Script:UserHomeDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13){
                        fcn_AddLogEntry ("... .. EA13 for $company is not set compare using FCT email")
                        If ($FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13 -eq $email){
                            fcn_AddLogEntry ("... .. EA13 on FA tenant is "+$FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13)
                            fcn_AddLogEntry ("... .. EA13 (email) from $company tenant is "+$email)
                            $Script:EA13 = $email
                        }
                        Else{
                            fcn_AddLogEntry ("... .. EA13 Missing or incorrect")
                            fcn_AddLogEntry ("... .. EA13 on FA tenant is "+$FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13)
                            fcn_AddLogEntry ("... .. EA13 from $company tenant is "+$email)
                            $Script:EA13=$email
                            $UpdateUser=$true}
                    }
                    Else{
                        If ($FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13 -eq $Script:UserHomeDetail.onPremisesExtensionAttributes.extensionAttribute13){
                            fcn_AddLogEntry ("... .. EA13 on FA tenant is "+$FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13)
                            fcn_AddLogEntry ("... .. EA13 from $company tenant is "+$Script:UserHomeDetail.onPremisesExtensionAttributes.extensionAttribute13)
                        }
                        Else{    
                            fcn_AddLogEntry ("... .. EA13 Missing or incorrect")
                            fcn_AddLogEntry ("... .. EA13 on FA tenant is "+$FALookupDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13)
                            fcn_AddLogEntry ("... .. EA13 from $company tenant is "+$Script:UserHomeDetail.onPremisesExtensionAttributes.extensionAttribute13)
                            $Script:EA13=$Script:UserHomeDetail.onPremisesExtensionAttributes.extensionAttribute13
                            $UpdateUser=$true
                        }
                    }
                }
                            
                If(($null -eq $FALookupDetail.companyName ) `
                -or ($FALookupDetail.companyName -ne $company)){
                    $UpdateUser=$true
                    fcn_AddLogEntry ("... . Missing or incorrect company name on FA tenant, update "+$FALookupDetail.companyName)
                }
    
                If($UpdateUser){
                    fcn_AddLogEntry ("... . Call function to update user attributes")
                    fcn_UpdateGuestAttributes $FALookupDetail.id $Script:UserHomeDetail $company $email $FALookupDetail
                }
                Else{
                    fcn_AddLogEntry ("... . No update required")
                }
    
            }       #employee matches
        }
        Else{          
            fcn_AddLogEntry ("... . No user with email matching $email found on FA Tenant")
            fcn_AddLogEntry ("... . Next check FA Tenant to see if there is an employeeID match")
                
            #check here to see if we find match on EmployeeID, if we do then user was renamed employeeIDs are not reused.
            If(($null -eq $Script:UserHomeDetail.EmployeeId) -and ($Tenant -eq $HWTenant)){
                If(($SAM -like "TE-*") -or ($SAM -like "SY-*")){
                    #If home warranty and TE or SY users, no employee IDs only check using email.
                    fcn_AddLogEntry ("... . no match via email skipping employeeID check for HW for TE and SY")
                    fcn_AddGuestToFA $Script:UserHomeDetail $email $company 
                        $IsValid = $Script:results.IsValid
                        If($IsValid){$cntC++}Else{$cntError++}
                }
                Else{
                    #lookup using eagle ID
                    fcn_AddTenantUserLogEntry ("%%% FAHW User "+$email+" has no employeeID")
                    $EgleIDfilter = "`$filter=extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute12 eq `'$Script:EagleID'"
                    $lookupBetaURL = $GraphBetaUsersURL+'?'+$EgleIDfilter
                    Try{$Lookup = Invoke-WebRequest -UseBasicParsing -Headers $FAauthToken -Uri $lookupBetaURL -Method GET
                        [Array]$FALookupDetail = (ConvertFrom-Json -InputObject $Lookup.Content).value}
                    Catch{
                        fcn_AddErrorLogEntry ("### unknown error in get by EagleID for Republic Title - Skipping User")
                        $cntSkip++
                        Continue
                    }
                }
            }
            ElseIf(($null -eq $Script:UserHomeDetail.EmployeeId) -and ($Tenant -eq $FCTTenant)){
                #If home warranty and TE or SY users, no employee IDs only check using email.
                fcn_AddErrorLogEntry ("### FCT User "+$email+" has no employeeID")
                fcn_AddTenantUserLogEntry ("### FCT User "+$email+" has no employeeID")
                fcn_AddErrorLogEntry ("### Skipping User")
                    $cntSkip++
                    Continue
                #fcn_AddGuestToFA $Script:UserHomeDetail $email $company 
                #    $IsValid = $Script:results.IsValid
                #    If($IsValid){$cntC++}Else{$cntError++}
            }
            ElseIf(($null -eq $Script:UserHomeDetail.EmployeeId) -and ($Tenant -eq $RTTenant)){
                #missing employeeID check using EagleID
                fcn_AddTenantUserLogEntry ("%%% Republic Title User "+$email+" has no employeeID")
                fcn_AddTenantUserLogEntry ("%%% Skipping User")
                #$EgleIDfilter = "`$filter=extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute12 eq `'$Script:EagleID'"
                #$lookupBetaURL = $GraphBetaUsersURL+'?'+$EgleIDfilter
                #Try{$Lookup = Invoke-WebRequest -UseBasicParsing -Headers $FAauthToken -Uri $lookupBetaURL -Method GET
                #    [Array]$FALookupDetail = (ConvertFrom-Json -InputObject $Lookup.Content).value}
                #Catch{
                #    fcn_AddErrorLogEntry ("### unknown error in get by EagleID for Republic Title - Skipping User")
                #    $cntSkip++
                #    Continue
                #}
            }           
            ElseIf($FAGuests.EmployeeID -contains $Script:UserHomeDetail.EmployeeId){ 
                $tmpUser = $FAGuests | Where-Object {$_.EmployeeID -eq $Script:UserHomeDetail.EmployeeId}
                fcn_AddLogEntry ("... . eMail did not match but employee IDs did, lookup using employeeID")
                fcn_AddLogEntry ("... .. Home Tenant: "+ $Script:UserHomeDetail.EmployeeId)
                fcn_AddLogEntry ("... .. FA Tenant:   "+ $tmpUser.EmployeeId)
                fcn_AddLogEntry ("... .. FA Tenant mail:   "+ $tmpUser.mail)
                $FALookupDetail=$null; $lookup=$null 
                $EID = $Script:UserHomeDetail.EmployeeId
                $EIDfilter = "`$filter=employeeID eq `'$EID'"
                $lookupBetaURL = $GraphBetaUsersURL+'?'+$EIDfilter
                Try{$Lookup = Invoke-WebRequest -UseBasicParsing -Headers $FAauthToken -Uri $lookupBetaURL -Method GET
                    [Array]$FALookupDetail = (ConvertFrom-Json -InputObject $Lookup.Content).value}
                Catch{
                    fcn_AddErrorLogEntry ("### unknown error in get by user by employee id filter - Skipping User")
                    $cntSkip++
                    Continue
                }

                If($Script:UserHomeDetail.EmployeeId -ne $FALookupDetail.employeeid){
                    fcn_AddLogEntry ("... # Employee IDs do not match between tenants, error - skipping user")
                    continue}
                
                # Is user a member of any groups?
                $Script:Rename=$true  
                
                #commented out for now, need additional testing and review
                $Script:UserGrps=$null 
                fcn_GetGrpMembership $FALookupDetail
                $Script:UserGrps = $Script:results.FAGrps
                $Isvalid = $Script:results.IsValid
                If(!($IsValid)){
                    fcn_AddErrorLogEntry ("... # Unable to get current FA Group membership, manually rename")
                    fcn_AddLogEntry ("### # Skipping rename")
                    $cntSkip++
                    Continue}

                fcn_AddGuestToFA $Guest $email $company
                $Isvalid = $Script:results.IsValid
                #$rc = [Array]$Groups = $Script:results.RC 
                If(!($Isvalid)){
                #Create user failed, throw error code and continue
                    Continue}
                fcn_ApplyGrpMembership $Guest $Groups

            }
            Else{
                #match not found, create as new user
                fcn_AddLogEntry ("... . no match via email or employeeID, adding new guest to FA")
                fcn_AddGuestToFA $Script:UserHomeDetail $email $company 
                    $IsValid = $Script:results.IsValid
                If($IsValid){$cntC++}Else{$cntError++}
            }
        }   #### This is the end of the Else
        fcn_AddLogEntry ("... Finished User")
        
    }   # ForEach User

}       #End Function

Function fcn_CheckEmployeeID{
    Param($FAGuest, $tmpAuthToken)

    #Look up stale user on home tenant using employeeID might be a name change
    $lookupBetaURL=$null; $HomeLookup = $null; $HomeLookupDetail=$null; $EID=$null
    $EID = $FAGuest.EmployeeID
    If($null -eq $EID){
        fcn_AddErrorLogEntry ("... EmployeeID is blank, skip")
        Continue 
    }
    $EIDfilter = "`$filter=employeeID eq `'$EID'"
    $lookupBetaURL = $GraphBetaUsersURL+'?'+$EIDfilter
    
    Try{$homeLookup = Invoke-WebRequest -UseBasicParsing -Headers $RTauthToken -Uri $lookupBetaURL -Method GET
        [Array]$HomeLookupDetail = (ConvertFrom-Json -InputObject $HomeLookup.Content).value 

        fcn_AddErrorLogEntry ("... EmployeeID "+$EID+" match found, process name change")
        #fcn_NameChange $faguest $HomeLookupDetail
    }
    Catch{
        fcn_AddErrorLogEntry ("... EmployeeID "+$EID+" not found, remove user")
        #fcn_RemoveUser $faguest 
    }

}

Function fcn_CheckForStaleUsers{
    Param($Users, $FAGuests, $tmpAuthToken)

    ForEach($faguest in $FAGuests){

        $UserURL = $GraphBetaUsersURL+$faguest.ID
        $ValUser = (Invoke-WebRequest -UseBasicParsing -Headers $FAauthToken -Uri $UserURL -Method GET)
        $UserDetail = ConvertFrom-Json -InputObject $ValUser.Content

        $tmpEA13 = $UserDetail.extension_4d5db290c1824986815f308e8a5a1f09_extensionAttribute13
        $found=$null
        $Found = $Users | where-object {$_.mail -eq $tmpEA13}
        If($null -eq $Found){
            #stale user delete    
            fcn_AddLogEntry ("... . "+$UserDetail.displayname+" not found in home tenant using email of $tmpEA13")
            $EIDFound=$null
            $EIDFound = $Users | where-object {$_.employeeId -eq $UserDetail.employeeId}
            If($null -eq $EIDFound){
                fcn_AddLogEntry ("... .. No match found using employeeID of "+$UserDetail.employeeId+" stale user remove from FA Tenant")
            }
            Else{
                fcn_AddLogEntry ("... .. Found Employee ID match: "+$UserDetail.employeeId+" must be a rename")
            }
        }
        Else{
            #Found email address check employee ID or EagleID
            
            If($UserDetail.employeeId -eq $Found.employeeid){
                #Write-host ("Same user employeeIDs match "+$Found.employee)
            }
            Else{
                #either user doesn't have employeeID or account name was reused
                fcn_AddLogEntry ("... % Found "+$UserDetail.displayname+" with email of $tmpEA13 but Employee IDs do not match")
                fcn_AddLogEntry ("... .. FA Employee ID:  "+$UserDetail.employeeId)
                fcn_AddLogEntry ("... .. $Tenant Employee ID: "+$Found.employeeId)
            
            }
        
        }



        If($Users.mail -contains $faGuest.Mail){
            #fcn_AddLogEntry ("... Verified as active User: "+$faGuest.Mail)
        }
        Else{
            fcn_AddLogEntry ("... User stale, is this a name change: "+$faGuest.displayname)
            fcn_CheckEmployeeID $FAGuest $tmpAuthToken
            $IsValid = $Script:results.IsValid
            If($IsValid){$cntR++}Else{$cntError++}
        } 
    }
}

######################################################################################################################################
######################################################################################################################################
#
#   Main Section
#
######################################################################################################################################
######################################################################################################################################

Write-host "... "

fcn_AddErrorLogEntry ("==================================================================")

fcn_AddLogEntry "..."
fcn_AddErrorLogEntry "... Script Starting"
If($Script:TestOnly){
    fcn_AddErrorLogEntry "***"
    fcn_AddErrorLogEntry "*** Test Only flag is set no errors should be logged only warnings"
    fcn_AddErrorLogEntry "***"
}

################################################################################################################################
# 2.0 Process the Home Warranty Users to invite them as Guests
################################################################################################################################

fcn_AddLogEntry "..."
fcn_AddLogEntry "... ------------------------------"
fcn_AddErrorLogEntry "... Starting Home Warranty"
fcn_AddLogEntry "... ------------------------------"

##############################################################################
# 2.1 This is where we authenticate and get our access token for Home Warrenty
##############################################################################
$Error.clear()
$Script:HomeTenantLog = $Script:MissingAttribFAHW
fcn_AddTenantUserLogEntry "... Auth to Home Warranty Tenant"

$HWBody       = @{grant_type="client_credentials";resource=$resource;client_id=$HWappID;client_secret=$HWSecret}
Try{$HWOauth      = Invoke-RestMethod -Method POST -Uri $loginURL/$HWTenant/oauth2/token?api-version=1.0 -Body $HWBody}
Catch{
    fcn_AddErrorLogEntry "##################################################"
    fcn_AddErrorLogEntry "### Auth for Home Warranty failed.. stopping ###"
    fcn_AddErrorLogEntry ("### # "+$Error.Exception+" ")
    fcn_AddErrorLogEntry "### "
    fcn_AddErrorLogEntry "### Expired Secret need updated Home Warranty Credentials ###"
    fcn_AddErrorLogEntry "##################################################"
    fcn_SendMessage
    $AuthOK=$false
}
if ($null -ne $HWOauth.access_token){
    $HWauthToken = @{'Authorization'="$($HWOauth.token_type) $($HWOauth.access_token)"}
    $AuthOK=$true
}

##############################################################################
# 2.2 Authenticate to FiratAm tenant
##############################################################################
fcn_AddLogEntry "... Auth to FA Tenant"
$Error.clear()
$FAbody       = @{grant_type="client_credentials";resource=$resource;client_id=$FAappID;client_secret=$FASecret}
Try{$FAOauth      = Invoke-RestMethod -Method POST -Uri $loginURL/$FATenant/oauth2/token?api-version=1.0 -Body $FAbody}
Catch{
    fcn_AddErrorLogEntry "##################################################"
    fcn_AddErrorLogEntry "### Auth for FA failed.. stopping ###"
    fcn_AddErrorLogEntry ("### # "+$Error.Exception+" ")
    fcn_AddErrorLogEntry "### "
    fcn_AddErrorLogEntry "### Failure need updated FA Credentials ###"
    fcn_AddErrorLogEntry "##################################################"
}
if ($null -ne $FAOauth.access_token){
    $FAauthToken = @{'Authorization'="$($FAOauth.token_type) $($FAOauth.access_token)"}
}

##############################################################################
# 2.3 Get the members of the Home Warranty control group
##############################################################################
If($AuthOK){
    $Entry = ("... get the Home Warranty Control Group $HWCtrlGrpName from Home Warranty Tenant")
    fcn_AddLogEntry $Entry
    
    #Call function which gets group and returns members
    fcn_GetAzureGroup $HWauthToken  $HWCtrlGrpGraphURL $HWCtrlGrpMembersGraphURL $HWCtrlGrpName $HWTenant
        $IsValid = $Script:results.IsValid
        $HWUsers = $Script:results.GrpMembers
    If(!($Isvalid)){
        fcn_AddErrorLogEntry "### Error occurred getting list from FAHW of accounts to create as Guests on FA tenant ... Skipping Home Warranty ###"
        $CritError++
        fcn_SendMessage
        Return
    }
    fcn_AddErrorLogEntry ("... Home Warranty Group has "+$HWUsers.count+" members")
    fcn_AddLogEntry ("... ")

    ##############################################################################
    # 2.4 Get list of HW users already in FA Tenant
    ##############################################################################

    [Array]$FAGuests=@()
    fcn_AddLogEntry ("... Get the Home Warranty control group from FA Tenant")
    fcn_GetAzureGroup $FAauthToken  $FAtoHWGuestsGrpGraphURL $FAtoHWGuestsMembersURL $FAListofHWGuests_GrpName $FATenant
        $IsValid = $Script:results.IsValid
        $FAGuests = $Script:results.GrpMembers

    If(!($Isvalid)){
        fcn_AddLogEntry "### Error occurred getting list of Home Warranty FA Guests, unable to validate terminations or disabled accounts ... Skipping ###"
    }
    Else{
        fcn_AddLogEntry ("... Home Warranty FA Guest group has "+$FAGuests.count+" members in FA Tenant")
        fcn_AddLogEntry "... "           
    }

    ##############################################################################
    # 2.5 Cycle through the list to create new guests
    ##############################################################################
    fcn_AddLogEntry ("... Cycle through the Home Warranty Users")
    fcn_AddLogEntry ("... ")
    fcn_ProcessUsers $HWUsers $HWTenant $HWauthToken

    fcn_AddLogEntry "... Done checking Guest users from Home Warranty, now look for stale guests accounts"

    ##############################################################################
    # 2.6 Perform CRUD activities for HW Guests
    ##############################################################################
    If(!($Isvalid)){
        #fcn_AddLogEntry "### Error occurred getting list of Home Warranty FA Guests, unable to validate terminations or disabled accounts ... Skipping ###"
    }
    Else{
        fcn_AddLogEntry ("... Check Home Warranty Users for CRUD Activities - skip for now")
        fcn_AddLogEntry "... "
        #fcn_CheckForStaleUsers $HWUsers $FAGuests $HWauthToken
    }

    fcn_AddErrorLogEntry "... "
    fcn_AddLogEntry "... ------------------------------"
    fcn_AddErrorLogEntry "... Finished Home Warranty"
    fcn_AddLogEntry "... ------------------------------"
    fcn_AddLogEntry "... "
}
Else{
    fcn_AddErrorLogEntry "### Skipped Processing Home Warranty"    
}

################################################################################################################################
#
# 3.0 Now start on Republic Title Users to create new guests
#
################################################################################################################################

fcn_AddLogEntry "... "
fcn_AddLogEntry "---------------------------------------"
fcn_AddErrorLogEntry "... Starting Republic Title"
fcn_AddLogEntry "---------------------------------------"

##############################################################################
# 3.1 This is where we authenticate and get our access token for Republic title
##############################################################################
$Script:HomeTenantLog = $Script:MissingAttribRT
fcn_AddTenantUserLogEntry "... Auth to Republic Title Tenant"
$Error.clear()

$RTBody       = @{grant_type="client_credentials";resource=$resource;client_id=$RTAppID;client_secret=$RTSecret}
Try{$RTOauth      = Invoke-RestMethod -Method POST -Uri $loginURL/$RTTenant/oauth2/token?api-version=1.0 -Body $RTBody}
Catch{
    fcn_AddLogEntry " "
    fcn_AddErrorLogEntry "##############################################################"
    fcn_AddErrorLogEntry "### Auth for Repulic Title failed.. STOPPING"
    fcn_AddErrorLogEntry ("### # "+$Error.Exception+" ")
    fcn_AddErrorLogEntry "### "
    fcn_AddErrorLogEntry "### Expired Secret need updated Republic Title Credentials ###"
    fcn_AddErrorLogEntry "##############################################################"
    fcn_SendMessage
    $AuthOK=$false
    #Return 
}
if ($null -ne $RTOauth.access_token){
    $RTAuthToken = @{'Authorization'="$($RTOauth.token_type) $($RTOauth.access_token)"}
    $AuthOK=$true
}

##############################################################################
# 3.2 Authenticate to FiratAm tenant
##############################################################################
#Authenticate to FA Destination location again so we don't time out
fcn_AddLogEntry "... Auth to FA Tenant second time"
$Error.clear()
$FAbody       = @{grant_type="client_credentials";resource=$resource;client_id=$FAappID;client_secret=$FASecret}
Try{$FAOauth      = Invoke-RestMethod -Method POST -Uri $loginURL/$FATenant/oauth2/token?api-version=1.0 -Body $FAbody}
Catch{
    fcn_AddErrorLogEntry "### Auth for FA failed.. stopping ###"
    fcn_AddErrorLogEntry ("### # "+$Error.Exception+" ")
    fcn_AddErrorLogEntry "### "
    fcn_AddErrorLogEntry "### Failure need updated FA Credentials ###"
    fcn_SendMessage
    Return 
}
if ($null -ne $FAOauth.access_token){
    $FAauthToken = @{'Authorization'="$($FAOauth.token_type) $($FAOauth.access_token)"}
}

##############################################################################
# 3.3 Get the members of the Republic Title control group
##############################################################################
If($AuthOK){
    $Entry = ("... get the Republic Title Control Group $RTCtrlGrpName from the Republic Title Tenant")
    fcn_AddLogEntry $Entry    

    #Call function which gets group and returns members
    [Array]$RTUsers=@()
    fcn_GetAzureGroup $RTAuthToken  $RTCtrlGrpGraphURL $RTCtrlGrpMembersGraphURL $RTCtrlGrpName $RTTenant
        $IsValid = $Script:results.IsValid
        $RTUsers = $Script:results.GrpMembers
    If(!($Isvalid)){
        ffcn_AddErrorLogEntry "### Error occurred getting list from Republic Title of accounts to create as Guests on FA tenant ... stopping ###"
        fcn_SendMessage
        
    }
    fcn_AddErrorLogEntry ("... Republic Title Group has "+$RTUsers.count+" members")

    ##############################################################################
    # 3.4 Get list of RT users already in FA Tenant
    ##############################################################################
    [Array]$FAGuests=@()
    fcn_AddLogEntry ("... Get the Republic Title control group from FA Tenant")
    fcn_GetAzureGroup $FAauthToken  $FAtoRTGuestsGrpGraphURL $FAtoRTGuestsMembersURL $FAListofRTGuests_GrpName $FATenant
        $IsValid = $Script:results.IsValid
        $FAGuests = $Script:results.GrpMembers

    If(!($Isvalid)){
        fcn_AddLogEntry "### Error occurred getting list of Republic Title FA Guests, unable to validate terminations or disabled accounts ... Skipping ###"
    }
    Else{
        fcn_AddLogEntry ("... Republic Title FA Guest group has "+$FAGuests.count+" members in FA Tenant")
        fcn_AddLogEntry "... "   
    }

    ##############################################################################
    # 3.5 Cycle through the list to create new guests
    ##############################################################################
    fcn_AddLogEntry ("... ")
    fcn_AddLogEntry ("... Cycle through the Republic Title Users")
    fcn_ProcessUsers $RTUsers $RTTenant $RTAuthToken
    fcn_AddLogEntry ("... ")
    fcn_AddLogEntry ("... Done checking users from Republic Title, now look for stale guests")


    ##############################################################################
    # 3.6 Perform CRUD activities for Republic Title Guests
    ##############################################################################
    [Array]$FAGuests=@()
    fcn_AddLogEntry ("... Get the Republic Title Warranty control group from FA Tenant")
    fcn_GetAzureGroup $FAauthToken  $FAtoRTGuestsGrpGraphURL $FAtoRTGuestsMembersURL $FAListofRTGuests_GrpName $FATenant
        $IsValid = $Script:results.IsValid
        $FAGuests = $Script:results.GrpMembers
    If(!($Isvalid)){
        fcn_AddLogEntry "### Error occurred getting list of Republic Title FA Guests, unable to validate terminations or disabled accounts ... stopping ###"
    }
    Else{
        fcn_AddLogEntry ("... Republic Title FA Guest group has "+$FAGuests.count+" members")
        fcn_AddLogEntry ("... For now skip check for stale users")
        #fcn_CheckForStaleUsers $RTUsers $FAGuests $RTAuthToken
    }

    fcn_AddLogEntry "... ------------------------------"
    fcn_AddErrorLogEntry "... Finished Republic Title"
    fcn_AddLogEntry "... ------------------------------"

}
Else{
    fcn_AddErrorLogEntry "### Skipped Processing Republic Title"
}

################################################################################################################################
# 4.0 Now start on First Canadian Trust to create new guests
################################################################################################################################
#This is where we authenticate and get our access token for Republic Title

fcn_AddLogEntry "... "
fcn_AddErrorLogEntry "... Starting First Canadian Trust"
fcn_AddLogEntry "---------------------------------------"

##############################################################################
# 4.1 This is where we authenticate and get our access token for Republic title
##############################################################################
$Error.clear()
$Script:HomeTenantLog = $Script:MissingAttribCan
fcn_AddTenantUserLogEntry "... Auth to First Canadian Trust"

$FCTBody       = @{grant_type="client_credentials";resource=$resource;client_id=$FCTAppID;client_secret=$FCTSecret}
Try{$FCTOauth      = Invoke-RestMethod -Method POST -Uri $loginURL/$FCTTenant/oauth2/token?api-version=1.0 -Body $FCTBody}
Catch{
    fcn_AddErrorLogEntry "########################################################"
    fcn_AddErrorLogEntry "### Auth for First Canadian Trust failed.. stopping ###"
    fcn_AddErrorLogEntry ("### # "+$Error.Exception+" ")    
    fcn_AddErrorLogEntry "### "
    fcn_AddErrorLogEntry "### Failure need updated FCT Secret ###"
    fcn_AddErrorLogEntry "### Expired Secret need updated FCT Credentials      ###"
    fcn_AddErrorLogEntry "########################################################"
    fcn_SendMessage
    $AuthOK=$false 
}
if ($null -ne $FCTOauth.access_token){
    $FCTAuthToken = @{'Authorization'="$($FCTOauth.token_type) $($FCTOauth.access_token)"}
    $AuthOK=$true 
}

##############################################################################
# 4.2 Authenticate to FiratAm tenant
##############################################################################
#Authenticate to FA Destination location again so we don't time out
fcn_AddLogEntry "... Auth to FA Tenant third time"
$Error.clear()
$FAbody       = @{grant_type="client_credentials";resource=$resource;client_id=$FAappID;client_secret=$FASecret}
Try{$FAOauth      = Invoke-RestMethod -Method POST -Uri $loginURL/$FATenant/oauth2/token?api-version=1.0 -Body $FAbody}
Catch{
    fcn_AddErrorLogEntry "### Auth for FA failed.. stopping ###"
    fcn_AddErrorLogEntry ("### # "+$Error.Exception+" ")
    fcn_AddErrorLogEntry "### "
    fcn_AddErrorLogEntry "### Failure need updated FA Credentials ###"
    fcn_SendMessage
    Return 
}
if ($null -ne $FAOauth.access_token){
    $FAauthToken = @{'Authorization'="$($FAOauth.token_type) $($FAOauth.access_token)"}
}

##############################################################################
# 4.3 Get the members of the First Canadian Title control group
##############################################################################
If($AuthOK){
    $Entry = ("... get the First Canadian Title Control Group $FCTCtrlGrpName from the fctca.onmicrosoft.com Tenant")
    fcn_AddLogEntry $Entry
    #out-file -FilePath $Script:LogPath\$Script:MissingAttribRT -InputObject "$DateTime  $Entry" -Append

    #Call function which gets group and returns members
    [Array]$FCTUsers=@()
    fcn_GetAzureGroup $FCTAuthToken  $FCTCtrlGrpGraphURL $FCTCtrlGrpMembersGraphURL $FCTCtrlGrpName $FCTTenant
        $IsValid = $Script:results.IsValid
        $FCTUsers = $Script:results.GrpMembers
    If(!($Isvalid)){
        ffcn_AddErrorLogEntry "### Error occurred getting list from First Canadian Title of accounts to create as Guests on FA tenant ... stopping ###"
        fcn_SendMessage
        Return
    }

    fcn_AddErrorLogEntry ("... First Canadian Title Group has "+$FCTUsers.count+" members")

    ##############################################################################
    # 4.4 Get list of FCT users already in FA Tenant
    ##############################################################################

    [Array]$FAGuests=@()
    fcn_AddLogEntry ("... Get the First Canadian Title control group from FA Tenant")
    fcn_GetAzureGroup $FAauthToken  $FAtoFCTGuestsGrpGraphURL $FAtoFCTGuestsMembersURL $FAListofFCTGuests_GrpName $FATenant
        $IsValid = $Script:results.IsValid
        $FAGuests = $Script:results.GrpMembers

    If(!($Isvalid)){
        fcn_AddLogEntry "### Error occurred getting list of First Canadian Title FA Guests, unable to validate terminations or disabled accounts ... Skipping ###"
    }
    Else{
        fcn_AddLogEntry ("... First Canadian Title FA Guest group has "+$FAGuests.count+" members in FA Tenant")
        fcn_AddLogEntry "... "   
    }

    ##############################################################################
    # 4.5 Cycle through the list to create new guests
    ##############################################################################
    fcn_AddLogEntry ("... ")
    fcn_AddLogEntry ("... Cycle through the First Canadian Title Users")    
    fcn_ProcessUsers $FCTUsers $FCTTenant $FCTAuthToken 
    fcn_AddLogEntry ("... ")
    fcn_AddLogEntry ("... Done checking users from First Canadian Title, now look for stale guests")

    ##############################################################################
    # 4.6 Perform CRUD activities for FCT Guests
    ##############################################################################
    [Array]$FAGuests=@()
    fcn_AddLogEntry ("... Get the Republic Title Warranty control group from FA Tenant")
    fcn_GetAzureGroup $FAauthToken  $FAtoFCTGuestsGrpGraphURL $FAtoFCTGuestsMembersURL $FAListofFCTGuests_GrpNamee $FATenant
        $IsValid = $Script:results.IsValid
        $FAGuests = $Script:results.GrpMembers
    If(!($Isvalid)){
        fcn_AddLogEntry "### Error occurred getting list of FCT FA Guests in the FA Tenant, unable to validate terminations or disabled accounts ... stopping ###"
    }
    Else{
        fcn_AddLogEntry ("... Republic Title FA Guest group has "+$FAGuests.count+" members")
        fcn_AddLogEntry ("... For now skip check for stale users")
        #fcn_CheckForStaleUsers $FCTUsers $FAGuests $FCTAuthToken
    }
    
    fcn_AddErrorLogEntry "... "
    fcn_AddLogEntry "... ------------------------------"
    fcn_AddErrorLogEntry "... Finished First Canadian Title"
    fcn_AddLogEntry "... ------------------------------"
    fcn_AddLogEntry "... "
}
Else{
    fcn_AddErrorLogEntry "### Skipped Processing First Canadian Title"
}

##############################################################################
# Script completed
##############################################################################

#create control file
$day = get-date -format dddd
$datefile = $date+".txt"
$DateTime = Get-Date
out-file -FilePath $Script:LogPath\$datefile -InputObject $DateTime -Force

fcn_AddErrorLogEntry " "
fcn_AddErrorLogEntry "... script complete"
fcn_AddErrorLogEntry " "

