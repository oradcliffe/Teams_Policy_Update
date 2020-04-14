#Requires -Modules AzureAD,SkypeOnlineConnector
[CmdletBinding()]
param (
    [parameter (Mandatory, Position = 1)][string]$teamsURI
)
Write-Verbose "Starting script"
function Write-ToTeams {
    [CmdletBinding()]
    param (
        [parameter (Position = 1)][string]$uri,
        [parameter (Position = 2)][string]$title,
        [parameter (Position = 3)][string]$text,
        [parameter (Mandatory, Position = 4)][string]$message
    )
    $body = ConvertTo-Json -Depth 4 @{
        title = "$title"
        text = "$text"
        sections = @(
            @{
                activityTitle = 'Teams Meeting Policy Application'
                activitySubtitle = 'Teams Meeting Policy applied to users.'
                activityText = 'This is an automated message to alert that some users have a modified Teams Meeting policy.'
                activityImage = "https://hotemoji.com/images/dl/6/thumbs-up-emoji-by-google.png" # this value would be a path to a nice image you would like to display in notifications
            },
            @{
                title = 'Users modifed with new Teams Meeting Policy:'
                text = "$message"
            }
        )
    }
    Invoke-RestMethod -uri $uri -Method Post -body $body -ContentType 'application/json'
}
function Invoke-TeamsLicenseCheck {
    #Requires -Modules AzureAD

    <#
    .SYNOPSIS
        Enumerates the given group and checks for Teams licenses
    
    .DESCRIPTION
        Enumerates the given group and checks for Teams licenses and returns those who are licensed. 
    
    .PARAMETER group
        This is a string with the name of the group that you wish to apply a policy to. 
    .PARAMETER credentials
        This is the credentials for someone with access to read Azure AD in a PSCredential Object
    #>
    [CmdletBinding()]
    param(
        [parameter (Mandatory, Position = 1)][string]$group, 
        [parameter (Position = 2)][System.Management.Automation.PSCredential]$credentials
    )
    try {
        Write-Verbose "Checking if AzureAD module already connected"
        Get-AzureADTenantDetail -ErrorAction SilentlyContinue | Out-Null
    }
    catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException]{
        Write-Verbose "Connecting to AzureAD"
        Connect-AzureAD -Credential $credentials
    }
    Write-Verbose "Checking $group"
    $output = @()
    $groupObject = Get-AzureADGroup -SearchString $group
    $members = Get-AzureADGroupMember -ObjectId $groupObject.ObjectId -All $true
    $count = $members.count
    Write-Verbose $count
    $x=1
    foreach ($user in $members) {
        Write-Verbose "Group member $x of $count"
        Write-Verbose $user.UserPrincipalName
        $userObject = Get-AzureADUserLicenseDetail -ObjectId $user.ObjectId
        if ($userObject.ServicePlans | Where-Object {$_.ProvisioningStatus -eq "Success" -and $_.ServicePlanName -like "*Teams*"}) {
           $tempuser = New-Object PSObject
           $mailaddress = $user.UserPrincipalName
           $tempuser | Add-Member -MemberType NoteProperty -Name "UPN" -Value $mailaddress
           $output += $tempuser
        }
        $x++
    }
    Write-Output $output
}
function Invoke-TeamsPolicyApplication {
    #Requires -Modules SkypeOnlineConnector
    [CmdletBinding()]
    param (
        [parameter (Position = 1, Mandatory, ParameterSetName = "PSCredObject")][System.Management.Automation.PSCredential]$creds,
        [parameter (Position = 2, Mandatory)]$users,
        [parameter (Position = 3, Mandatory)]$policyToApply
    )
    $VerbosePreference = 'Continue'
    Write-Verbose "Starting to set $policyToApply on Teams"
    Import-Module SkypeOnlineConnector
    Write-Verbose "Creating Skype Online PS Session"
    $sfbSession = New-CsOnlineSession -Credential $creds
    Write-Verbose "Created Skype Online Session"
    Import-PSSession $sfbSession -AllowClobber -CommandName Get-CsOnlineUser,Grant-CsTeamsMeetingPolicy | Out-Null
    Write-Verbose "Imported Skype PS Session"
    $policyApplied = @()
    foreach ($user in $users) {
        $upn = $user.upn
        if (Get-CsOnlineUser -Identity $upn | Where-Object {$_.TeamsMeetingPolicy -eq $policyToApply}) {
            Write-Verbose "Teams Policy is already correct on $upn"
        }
        else {
            Write-Verbose "Setting Teams Meeting Policy $policyToApply on $upn"
            $temp = New-Object PSObject
            $temp | Add-Member -MemberType NoteProperty -Name User -Value $upn
            $temp | Add-Member -MemberType NoteProperty -Name Policy -Value $policyToApply
            [array]$policyApplied += $temp
            Grant-CsTeamsMeetingPolicy -Identity $upn -PolicyName $policyToApply
            Remove-Variable temp
            Remove-Variable upn
        }
    }
    Write-Output $policyApplied
    Remove-PSSession $sfbSession
}
#Get Teams Admin Credentials
$teamsAdminCreds = Get-AutomationPSCredential "teamsAdmin" -Verbose
$policy1Users = Invoke-TeamsLicenseCheck -group "TeamsMeetingPolicyGroup1" -credentials $teamsAdminCreds
$policy2Users = Invoke-TeamsLicenseCheck -group "TeamsMeetingPolicyGroup2" -credentials $teamsAdminCreds
$policy3Users = Invoke-TeamsLicenseCheck -group "TeamsMeetingPolicyGroup3" -credentials $teamsAdminCreds
$teamsChanges = Invoke-TeamsPolicyApplication -creds $teamsAdminCreds -users $policy1Users -policyToApply "testMeeting1" -Verbose
$teamsChanges += Invoke-TeamsPolicyApplication -creds $teamsAdminCreds -users $policy2Users -policyToApply "testMeeting2" -Verbose
$teamsChanges += Invoke-TeamsPolicyApplication -creds $teamsAdminCreds -users $policy3Users -policyToApply "testMeeting3" -Verbose
if (!$teamsChanges) {
    $teamsChanges = "No changes have been made on this run."
}
else {
    [string]$teamsChanges = $teamsChanges | ConvertTo-Html -Fragment
}
Write-Verbose $teamsChanges
Write-ToTeams -uri $teamsURI -title "Teams Meeting Policy Assignment Notification" -text "Teams Meeting Policy assignment task successful" -message $teamsChanges