#Script to check for future staff events and copy them to
#mailboxes created in past 24hrs.

#Revision Jul 29 2015:  Tom Cheang
#Fixing script so that Meeting Updates sent by Comms Team
#  are processed by items created by this script.  Solution is to copy
#  the CleanGlobalObjectId to the n00bs calendar.

#Revision Aug 3, 2015:  Tom Cheang
#Reworking script function.  Using Organizer's calendar to send output
#  meeting invitation.  Remove added attendees once invite sent.


#Revision Sep 14, 2016
# stamp n00b mailboxes on customattr3 with 'Invites', send to these
#mailboxes since we can no longer key off of msExchWhenMailboxCreated

#region Functions
########## Declare Functions  ###########
#Main function
function Add-EventsToNewEmployeeMailboxes {
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory=$true)]
    [int]$CreateDateAddDays,
    $CommsMbx,
    $Credential
  )
    #Line below no longer necessary.  No longer get mailboxes via whencreated
 #  $NewMailboxes = Get-NewMailboxes -CreateDateAddDays $CreateDateAddDays
  $NewMailboxes = Get-NewRehireConversionMailboxes

  if (!($NewMailboxes)) {
    $Now = Get-Date -Format g
    Write-Output "No n00b/rehire mailboxes found between yesterday and $Now."
    Break
  }

  # Let's adapt this script so it works with your comm team's mailbox
  $intcom = Get-Recipient -Identity $CommsMbx
  if ($intcom.RecipientTypeDetails -eq 'RemoteUserMailbox' -or
      $intcom.RecipientTypeDetails -eq 'RemoteSharedMailbox') {

    $splNewService = @{
      O365 = $true
      Credential = $Credential
    }
    $service = New-DelegatedEwsService @splNewService
    $Events = Get-DelegatedFutureMeetingsTostaff -service $service -CommsMbx $CommsMbx
  }
  #else revert to on-prem EWS connection
  else {
    $service = Get-ImpersonatedEwsService -Impersonate $CommsMbx
    $Events = Get-FutureMeetingsTostaff -service $service
  }

  if (!($Events)) {
    Write-Output "No future events for staff found on comm's calendar."
    Break
  }

  foreach ($e in $Events) {
    $splAttendeeChanges = @{
      'service' = $service
      'Id' = $e.Id
    }
    # Add Attendees then send update
    $NewMailboxes | Add-AttendeesToMeeting @splAttendeeChanges
    # Undo changes to event.  Send no updates
    $NewMailboxes | Undo-AttendeeChanges @splAttendeeChanges
  }
}

function New-DelegatedEwsService {
  [CmdletBinding()]
  [Alias('Get-DelegatedEwsService')]
  Param (
    [switch]$O365, $Credential
  )
  if ($O365) {
    $url = 'https://outlook.office365.com/EWS/Exchange.asmx'
    $cred = $Credential.GetNetworkCredential()
    $ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($Exchver)
    $service.Credentials = New-Object System.Net.NetworkCredential -ArgumentList $cred.UserName,
      $cred.Password, $cred.Domain
    $service.url = New-Object Uri($url)
  }
  else {
    $ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($Exchver)
    $service.UseDefaultCredentials = $true
    $service.AutodiscoverUrl($Impersonate)
  }
  Return $service
}

#Returns impersonated service object
function New-ImpersonatedEwsService {
  [CmdletBinding()]
  [Alias('Get-ImpersonatedEwsService')]
  Param (
    [Parameter(Mandatory=$true,
    ValueFromPipelineByPropertyName=$true,
    Position=0)]
    $Impersonate,
    [switch]$O365,
    [PSCredential]$Credential
  )
  if ($O365) {
    $ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchVer)

    $ArgumentList = ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress), $Impersonate
    $ImpUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList $ArgumentList
    $service.ImpersonatedUserId = $ImpUserId

    $url = 'https://outlook.office365.com/EWS/Exchange.asmx'
    if (! $Credential) {
      $Credential = (Get-Credential -UserName (
        $env:UserName + '@fb.com') -m 'Credentials for O365 admin access'
      ).GetNetworkCredential()
    }
    $service.Credentials = New-Object System.Net.NetworkCredential -ArgumentList $Credential.UserName,
      $Credential.Password, $Credential.Domain
    $service.url = New-Object Uri($url)
  }
  else {
    $ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchver)
    $ArgumentList = ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress),$Impersonate
    $ImpUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList $ArgumentList
    $service.ImpersonatedUserId = $ImpUserId

    $service.UseDefaultCredentials = $true
    $service.AutodiscoverUrl($Impersonate)
  }
  Return $service
}

# Search "int-com@fb.com" by ItemView for future events sent to staff
function Get-FutureMeetingsTostaff {
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory=$true,
    ValueFromPipelineByPropertyName=$true,
    Position=0)]
  $service
  )
  $start = Get-Date
  $attendee = "all_employees"
  $CalendarFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(
    $service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)

  $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView("50")
  $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection(
    [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)

  $Filter1 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring(
    [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::RequiredAttendees, $attendee)

  $Filter2 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo(
    [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, $start)

  $SearchFilter.Add($Filter1)
  $SearchFilter.Add($Filter2)

  $psPropSet = new-object Microsoft.Exchange.WebServices.Data.PropertySet(
    [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
  $ItemView.PropertySet = $psPropSet

  #FindItems will return Attendee lists
  $FindItems = $service.FindItems($CalendarFolder.Id, $SearchFilter, $ItemView)
  foreach ($f in $FindResults){
    $f.Load($psPropSet)
  }
  Return $FindItems
}

function Get-DelegatedFutureMeetingsTostaff {
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory=$true,
    ValueFromPipelineByPropertyName=$true,
    Position=0)]
  $service, $CommsMbx
  )
  $start = Get-Date
  $attendee = "staff"
  $CalendarFolderName = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar
  $CalendarFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($CalendarFolderName, $CommsMbx)
  $CalendarFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $CalendarFolderId)

  $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView("500")
  $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection(
    [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)

  #$Filter1 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring(

  #  [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::RequiredAttendees,$f $attendee)

  $Filter2 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo(
    [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, $start)

  #$SearchFilter.Add($Filter1)
  $SearchFilter.Add($Filter2)

  $psPropSet = new-object Microsoft.Exchange.WebServices.Data.PropertySet(
    [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
  $ItemView.PropertySet = $psPropSet

  #FindItems will return Attendee lists
  $FindItems = $service.FindItems($CalendarFolder.Id, $SearchFilter, $ItemView)

  #load properties for each item in $finditems
  $service.LoadPropertiesForItems($FindItems, $psPropSet) | Out-Null

  Return $FindItems.Items | where {$_.RequiredAttendees.name -contains $attendee}
}

#Returns array of PrimarySmtpAddresses of new mailboxes
function Get-NewMailboxes {
  [CmdletBinding()]
  Param (
  [Parameter(Mandatory=$true,
     ValueFromPipelineByPropertyName=$true,
     Position=0)]
  [int]$CreateDateAddDays
  )
  $createDate = (Get-Date).AddDays("$CreateDateAddDays")

  $adSplat = @{
    Properties = @('proxyAddresses', 'msExchWhenMailboxCreated',
      'msExchExtensionCustomAttribute3')
    Filter = {msExchWhenMailboxCreated -gt $createDate}
  }
  $NewMailboxes = get-aduser @adSplat | select *,
    @{n='PrimarySmtpAddress';e={$_.mail}}
  if ($NewMailboxes) {
  Write-host "New Mailboxes found: $($Newmailboxes.PrimarySmtpAddress)"
  }
  Else {write-host "No new mailboxes found."}
  Return $NewMailboxes
}

function Get-NewRehireConversionMailboxes {
  [CmdletBinding()]

  $adSplat = @{
    SearchScope = 'Subtree'
    Properties = @('mail', 'msExchWhenMailboxCreated',
      'msExchExtensionCustomAttribute3')
    Filter = {msExchExtensionCustomAttribute3 -like '*'}
  }
  $Rehires = get-aduser @adSplat | where {
    $_.msExchExtensionCustomAttribute3 -contains 'Rehire-Conversion' -OR
    $_.msExchExtensionCustomAttribute3 -contains 'Invites'} |
    select *, @{n='PrimarySmtpAddress';e={$_.mail}}
  if ($Rehires) {
    Write-host "New/Rehire/Conv Mailboxes found:$($Rehires.PrimarySmtpAddress)"
  }
  Else {write-host "No rehire/conversion mailboxes found."}
  Return $Rehires
}

function Get-BindMeetingById {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$true,
    ValueFromPipelineByPropertyName=$true,
    Position=0)]
    $service, $Id
  )
  #Load up FirstClass ProperySet
  $psPropSet = new-object Microsoft.Exchange.WebServices.Data.PropertySet(
    [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
  $ItemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId($Id.UniqueId)
  $MatchingMeeting = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service,$ItemId, $psPropSet)

  Return $MatchingMeeting
}

function Add-AttendeesToMeeting {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$true,
      ValueFromPipelineByPropertyName=$true,
      Position=0)]
      $service,
    [Parameter(Mandatory=$true,
      ValueFromPipelineByPropertyName=$true)]
      $Id,
    [Parameter(Mandatory=$true,
      ValueFromPipelineByPropertyName=$true)]
      $PrimarySmtpAddress
  )

  Begin {
    #Let's bind to the mail message by Id
    $Invitation = Get-BindMeetingById -service $service -Id $Id
    Write-Output "Meeting: `"$($Invitation.Subject)`" at  $($Invitation.Start):"
    Write-Output "Sending invitations to the following n00bs:"
  }
  Process {
    $Invitation.RequiredAttendees.Add($PrimarySmtpAddress) | Out-Null
    Write-Output "          $($PrimarySmtpAddress)"
  }
  End {
    #Let's add the new attendees and send invitations only to new members
    $ConflictResolutionMode = [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite
    $SendUpdateMode = [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendOnlyToChanged

    $Invitation.Update($ConflictResolutionMode, $SendUpdateMode)
    Write-Host "***********************************************************"
    Write-Host "Invite sent to newly added attendees"
  }
}

function Undo-AttendeeChanges {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$true,
      ValueFromPipelineByPropertyName=$true,
      Position=0)]
      $service,
    [Parameter(Mandatory=$true,
      ValueFromPipelineByPropertyName=$true)]
      $Id,
    [Parameter(Mandatory=$true,
      ValueFromPipelineByPropertyName=$true)]
      $PrimarySmtpAddress
  )

  Begin {
    #Let's bind to the mail message by I
    $Meeting = Get-BindMeetingById -service $service -Id $Id
  }
  Process {
    for ($i=0; $i -lt $Meeting.RequiredAttendees.count; $i++) {
      if ($Meeting.RequiredAttendees[$i].Address -eq $PrimarySmtpAddress) {
        $Meeting.RequiredAttendees.RemoveAt($i)
      }
    }
  }
  End {
    #Let's now restore meeting to prior state.
    #Remove added attendees and send no updates.
    $ConflictResolutionMode = [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite
    $SendUpdateMode = [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToNone

    $Meeting.Update($ConflictResolutionMode, $SendUpdateMode)
    Write-Host "Cleanup - Invite restored to original RequiredAttendees list"
    Write-Host "***********************************************************`r`n"
  }
}

#endregion

#region Setup Env

#Load Exchange on-prem cmdlets
Enable-EMS -ShowTextOutput $false

#Let's try to load the EWS API library, and inform user.
if (!(Get-Module -Name 'Microsoft.Exchange.WebServices')) {
  $Module = "C:\Program Files\Microsoft\Exchange\" +
    "Web Services\2.2\Microsoft.Exchange.WebServices.dll"
  if (Test-Path $Module) {
    Import-Module -Name $Module
    Write-Output "EWS API module is loaded."
  }
  else {
    Write-Warning "EWS API not loaded.  Cannot find file: `r`n $Module"
  }
}

$CreateDateAddDays = '-1' # (in days)
$Now = Get-Date -Format g
#Mailbox containing staff events
$CommsMbx = "comms@yourcompany.com"

  #splat for main function
  $splMainFx = @{
    "CommsMbx" = $CommsMbx
    "CreateDateAddDays" = $CreateDateAddDays
  }


#endregion


##################### Begin script #####################

Add-EventsToNewEmployeeMailboxes @splMainFx

