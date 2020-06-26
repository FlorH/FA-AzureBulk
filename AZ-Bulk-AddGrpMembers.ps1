
$loginURL   = "https://login.microsoftonline.com"
$resource   = "https://graph.microsoft.com"
$GraphSPURL="https://myapps.microsoft.com/Beta/servicePrincipals"
$GraphAppsURL      = "https://graph.microsoft.com/Beta/Applications"
$GraphSPURL      = "https://graph.microsoft.com/Beta/ServicePrincipals"
$GraphSchema      = "https://graph.microsoft.com/Beta/SchemaExtensions"
$GraphGrpURL = "https://graph.microsoft.com/v1.0/groups"
$GraphGrpBURL = "https://graph.microsoft.com/Beta/groups/"
$GraphUsersURL      = "https://graph.microsoft.com/v1.0/Users"
$GraphUsersBURL      = "https://graph.microsoft.com/Beta/Users"

$Script:FATenant   = "firstam.onmicrosoft.com"                        
$FAappID    = "7b5fc675-995d-4a96-8572-84e1884f5b8b"           #SP-FAGuestUserInvite - this is the application we are using to call the graph API 
$FASecret   = "uP/+B7c/q0XdFFzKjGWcxCR6ok9aiLUWjxz2OE24D1c="    #this is the secret associated with the app
#

$Date = Get-Date
$Script:ReportPath="C:\PSReports"
$LogFile = "C:\PSReports\PSLogs\AZUpdateGrps-"+$date.tostring("MM")+"-"+$date.day+"-"+$date.year+".txt"
out-file -FilePath $SkipLog -InputObject '"Name","DisplayName","Email"' -Force

# this is where we get the group
$grp = "AAD-SG-TABLEAU-USERS"
$List = Import-csv C:\code\data\TableauUserList.csv

##################################################################################################
# Functions
##################################################################################################
function fcn_AddErrorLogEntry {
    param($Entry)
	$DateTime = Get-Date -format "yyyy/MM/dd HH:mm"
	Write-Host "$DateTime  $Entry"
    out-file -FilePath $LogFile -InputObject $Entry -Append
}

function fcn_AddLogEntry {
    param($Entry)
	$DateTime = Get-Date -format "yyyy/MM/dd HH:mm"
	Write-Host "$DateTime  $Entry"
    out-file -FilePath $LogFile -InputObject "$DateTime  $Entry" -Append
}

Function fcn_Auth{
	$Error.clear()
	$FAbody       = @{grant_type="client_credentials";resource=$resource;client_id=$FAappID;client_secret=$FASecret}
	Try{$FAOauth      = Invoke-RestMethod -Method POST -Uri $loginURL/$FATenant/oauth2/token?api-version=1.0 -Body $FAbody}
	Catch{
		fcn_AddErrorLogEntry "### Auth for FA failed.. stopping ###"
		fcn_AddErrorLogEntry ("### # "+$Error.Exception+" ")
		[Hashtable]$Script:results = @{isValid=$false; Message="# Failed to Authenticate #"}
		#Return 
	}
	if ($null -ne $FAOauth.access_token){
        $Script:FAauthToken = @{'Authorization'="$($FAOauth.token_type) $($FAOauth.access_token)"}
        fcn_AddLogEntry ("... Token granted")
	}
}

Function fcn_GetGrpMembers{
    Param($GrpList)
    
    $LookupGrpURL = $GraphGrpURL+"/"+$GrpList.ID+"/Members"
    
    Try{[Array]$MList = (Invoke-WebRequest -UseBasicParsing -Headers $Script:FAauthToken -Uri $LookupGrpURL -Method GET)}
    Catch{fcn_AddErrorLogEntry ("%%% % Unable to get Group members to build backup file")
        Continue}
    
    $Members = (ConvertFrom-Json -InputObject $Mlist.Content).Value
    $MemFile = $Script:ReportPath+"\Backup\"+$GrpList.DisplayName+"-Members-"+$date.tostring("MM")+"-"+$date.day+"-"+$date.year+".txt"
    If(Test-Path $MemFile){
        $Entry = "################################################## "
        out-file -FilePath $MemFile -InputObject $Entry -Append
        $Entry = " "
        out-file -FilePath $MemFile -InputObject $Entry -Append
        $Entry = "      Additional run of the script   $DateTime"
        out-file -FilePath $MemFile -InputObject $Entry -Append
        $Entry = " "
        out-file -FilePath $MemFile -InputObject $Entry -Append
        $Entry = "################################################## "
        out-file -FilePath $MemFile -InputObject $Entry -Append}
    Else{
        $Entry = "DisplayName,UPN,ID"
        out-file -FilePath $MemFile -InputObject $Entry -force}

    ForEach($Member in $Members){
        $Entry = '"'+$member.DisplayName+'","'+$Member.userPrincipalName+'","'+$member.ID+'"'
        out-file -FilePath $MemFile -InputObject $Entry -Append
    }

}

Function fcn_UserCycleMail{
    Param($UserDetail, $LookupMail, $lookupDisplayName)

    Write-Host "... check using email address $lookupmail" -ForegroundColor Yellow 

    [Array]$FoundUsers=$Null; $IsFound=$false 
    ForEach($User in $UserDetail){
        If($User.Mail -eq $LookupMail){
            $IsFound=$true
            $FoundUsers+=$User}
    }
    
    If($IsFound){
        Write-Host ("... found "+$FoundUsers.count+" match")
         If($FoundUsers.count -gt 1){
            Write-Host "... still have more than 1 matches"            
         }
         Else{
            Write-Host "... found single match using email"}
    }

    [Hashtable]$Script:results = @{IsFound=$IsFound; FoundUsers=$FoundUsers}
}

Function fcn_UserCycleUPN{
    Param($UserDetail, $lookupUPN, $LookupMail, $lookupDisplayName)

    $IsFound=$False; [Array]$FoundUsers=$null
    ForEach($User in $UserDetail){
        
        fcn_AddLogEntry ("... found match is "+ $User.userPrincipalName) 

        If($User.userType -eq "Guest"){
            If($User.companyName -eq "Republic Title"){
                $tmpUPN = $user.userPrincipalName.replace("_republictitle.com#EXT#@firstam.onmicrosoft.com","")
                fcn_AddLogEntry ("... . User is a Guest remove external reference on UPN")
                $IsFound=$true
            }
            If($User.companyName -eq "FAHW"){
                $tmpUPN = $user.userPrincipalName.replace("_fahw.com#EXT#@firstam.onmicrosoft.com","")
                fcn_AddLogEntry ("... . User is a Guest remove external reference on UPN")
                $IsFound=$true
            }
        }
        Else{$tmpUPN=$user.userPrincipalName.split("@")[0]}

        fcn_AddLogEntry ("... . use $tmpUPN as the compare to file name $lookupUPN")

        If($tmpUPN -eq $lookupUPN){
            $IsFound=$true
            $FoundUsers+=$User
            fcn_AddLogEntry ("... . found match using UPN for user "+$User.DisplayName)      
        }
        Else{fcn_AddLogEntry ("... . no match with $tmpUPN")}
    }

    If($IsFound){
        
        If($FoundUsers.count -gt 1){
            fcn_AddLogEntry ("... . still have more than 1 match, check using mail")
            fcn_UserCycleMail $FoundUsers $LookupMail $lookupDisplayName
                $IsFound = $Script:results.IsFound
                $FoundUsers = $Script:results.FoundUsers

        }
        #Else{fcn_AddLogEntry ("... . single match fround "+$FoundUsers.DisplayName)}
    }

    [Hashtable]$Script:results = @{IsFound=$IsFound; FoundUsers=$FoundUsers}
}


Function fcn_AddUsertoGroup{
    Param($FoundUsers)

    $NewMember=$null; $error.clear()
    $FoundUserURL = "https://graph.microsoft.com/beta/users/"+$FoundUsers.ID
    $JsonMember = @{
        "@odata.id" = $FoundUserURL
    } | ConvertTo-Json

    $AddUserURL = $GraphGrpURL+"/"+$GrpList.ID+"/members/`$ref"
    Try{$NewMember = Invoke-WebRequest -UseBasicParsing -Headers $FAauthToken -Uri $AddUserURL -Method POST -body $JsonMember -ContentType "application/json"}
    Catch{fcn_AddErrorLogEntry ("%%% % Unable to add user to group update skip file")
        $RC = (convertfrom-json $Error.errordetails.message).error}
    
    If($NewMember.statuscode -eq 204){
       fcn_AddLogEntry ("... . POST returned status code: "+$NewMember.statuscode+" user added")
       $status="add"       
    }
    ElseIf($RC.message -like "*One or more added object references already exist for the following modified properties: 'members'.*"){
        fcn_AddLogEntry ("...  % User is already a member of the group")
        $status="skip"
    }
    Else{
        fcn_AddLogEntry ("%%% % Unable to add "+ $FoundUsers.displayname +" to group, adding to skip file")
        $Entry = '"'+$lookupUPN+'","'+$lookupDisplayName+'","'+$lookupMail+'"'
        out-file -FilePath $SkipLog -InputObject $Entry -Append
        $status="failed"        
    }

    [Hashtable]$Script:results = @{Status=$Status}       
}

##################################################################################################
##################################################################################################
# Main Code
##################################################################################################
##################################################################################################
fcn_AddLogEntry ("...........................................................")
fcn_AddLogEntry ("... Starting Script")
fcn_AddLogEntry ("... Script run by "+$env:USERNAME)
fcn_AddLogEntry ("... ")

#authenticate to production tenant
FCN_AUTH

fcn_AddLogEntry ("... Lookup Group $grp")
$Grpfilter = "`$filter=(startswith(displayname,`'$Grp'"
$tmpGrpURL = $GraphGrpURL+'?'+$Grpfilter+"))"

Try{[Array]$Group = (Invoke-WebRequest -UseBasicParsing -Headers $Script:FAauthToken -Uri $tmpGrpURL -Method GET)}
Catch{
    fcn_AddErrorLogEntry "### Unable to find Group $Grp .... stopping ###"
        Return}

$GrpList = (ConvertFrom-Json -InputObject $Group.Content).Value
If($GrpList.count -ne 1){
    fcn_AddErrorLogEntry ("### More than one group returned, check group name .... stopping ###")
    Return
}

fcn_AddLogEntry ("... Group found")
fcn_AddLogEntry ("... Group Detail: "+$GrpList.Displayname+" Desc: "+$GrpList.Description+" ObjectID: "+$GrpList.ObjectID)
fcn_AddLogEntry ("... ")

fcn_AddLogEntry ("... Dump group members into a file")
fcn_GetGrpMembers $GrpList
fcn_AddLogEntry ("... ")
fcn_AddLogEntry ("... Found "+$List.count+" items in the input file")
fcn_AddLogEntry ("... ")

$addcnt=0; $skipcnt=0; $Failcnt=0
ForEach($item in $List){

    $lookupURL = $null; $LookupUPN=$null; $LookupMail=$null; $lookupDisplayName=$null 
    $lookupUPN = $item.name 
    $lookupMail = $item.email
    $lookupDisplayName = $item.DisplayName

    Write-Host ""
    fcn_AddLogEntry ("... Lookup $lookupUPN")
    #$mailfilter = "`$filter=mail eq `'$Script:InviteMail'"
 
    $UPNfilter = "`$filter=(startswith(UserPrincipalName,`'$lookupUPN'"
    $lookupURL = $GraphUsersBURL+'?'+$UPNfilter+"))"

    $TmpUser = (Invoke-WebRequest -UseBasicParsing -Headers $Script:FAauthToken -Uri $lookupURL -Method GET)
    [Array]$UserDetail = (ConvertFrom-Json -InputObject $tmpUser.Content).Value
   
    If($UserDetail.count -eq 0){
        fcn_AddLogEntry ("... Match not found, adding to skip file")
        $Entry = '"'+$lookupUPN+'","'+$lookupDisplayName+'","'+$lookupMail+'"'
        out-file -FilePath $SkipLog -InputObject $Entry -Append
        $Failcnt++     
            Continue}

    # userPrincipalName": "aarmiller_republictitle.com#EXT#@firstam.onmicrosoft.com"
    fcn_AddLogEntry ("... Found "+$UserDetail.count+" matches for $lookupUPN using UPN")
    If($UserDetail.count -eq 1){
        $FoundUsers = $Userdetail
        $IsFound=$true
    }
    Else{
        fcn_UserCycleUPN $UserDetail $lookupUPN $LookupMail $lookupDisplayName
            $IsFound = $Script:results.IsFound
            $FoundUsers = $Script:results.FoundUsers           
    }

    If($IsFound){
        If($FoundUsers.count -eq 1){
            fcn_AddLogEntry ("... Single match found for "+$FoundUsers.DisplayName+" using "+$FoundUsers.UserPrincipalName)
            fcn_AddLogEntry (".......... this is where we add the user to the group")
            fcn_AddUsertoGroup $FoundUsers
                $Status = $Script:results.Status
            If($Status -eq "add"){$addcnt++}
            ElseIf($Status -eq "skip"){$skipcnt++}
            Else{$failcnt++}            
        }
        Else{
            fcn_AddLogEntry ("... Still have more than one match found, add to skip file")
            $Entry = '"'+$lookupUPN+'","'+$lookupDisplayName+'","'+$lookupMail+'"'
            out-file -FilePath $SkipLog -InputObject $Entry -Append 
            $Failcnt++
        }    
    }
    Else{
        fcn_AddLogEntry ("... $LookupUPN not found in Azure, add to skip file")
        $Entry = '"'+$lookupUPN+'","'+$lookupDisplayName+'","'+$lookupMail+'"'
        out-file -FilePath $SkipLog -InputObject $Entry -Append   
        $Failcnt++     
    }
    fcn_AddLogEntry ("...........................................................")
}

fcn_AddLogEntry ("...  ")
fcn_AddLogEntry ("... Input file had "+$list.count+" total items")
fcn_AddLogEntry ("... "+$addcnt+" were added, and "+$skipcnt+" were already members of the group")
fcn_AddLogEntry ("... "+$failedcnt+" were not added and were plaaced in the skip file.")
fcn_AddLogEntry ("... Script finished")
fcn_AddLogEntry ("...  ")
Write-Host " "
