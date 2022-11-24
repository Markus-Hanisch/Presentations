#region Parameters
$projectPath = "$env:USERPROFILE\Documents\EXUSG\Set-Rooms\Set-RoomMailbox_Set-Place_Type-Workspace\Create-Set-Room-in-Bulk"
$csvName = 'EXUSG_AT_Rooms-Places-Workspaces.csv'
#endregion


#region Set Project Location
Set-Location $projectPath
#endregion

Connect-ExchangeOnline

Get-EXOMailbox -ResultSize Unlimited -Properties Identity, Alias, RecipientTypeDetails -Filter "RecipientTypeDetails -eq 'RoomMailbox'" | 
    Sort-Object -Property Identity | 
    Select-Object Identity, Alias, RecipientTypeDetails 

#region New-Mailbox
Import-Module -Name ExchangeOnlineManagement
# see also: https://docs.pexip.com/support/exchange_mailbox_ps.htm
Import-Csv -Path $csvName | ForEach-Object { 
    $RemotePowerShellEnabled = [System.Convert]::ToBoolean($_.RemotePowerShellEnabled)    
    
    # New-Mailbox: https://docs.microsoft.com/en-us/powershell/module/exchange/new-mailbox?view=exchange-ps
    New-Mailbox -Name $_.Name -Room `
        -Alias $_.Alias `
        -DisplayName $_.DisplayName `
        -RemotePowerShellEnabled $RemotePowerShellEnabled
}
#endregion


#region Set-Mailbox
Import-Csv -Path $csvName | ForEach-Object { 
    # Set-Mailbox: https://docs.microsoft.com/en-us/powershell/module/exchange/set-mailbox?view=exchange-ps
    $paramSetMailbox = @{
        AccountDisabled          = $true
        EnableRoomMailboxAccount = $false
        ResourceCapacity         = $_.Capacity
    }
    Set-Mailbox -Identity $_.Identity @paramSetMailbox

    $paramSetMailboxType = @{ 
        Type = $_.Type
    }
    Set-Mailbox -Identity $_.Identity @paramSetMailboxType
}
#endregion


#region Set-Mailbox - MailTip_only
$csvNameMailTipOnly = 'EXUSG_Rooms-Places-Workspaces_MailTip_only.csv'

Import-Csv -Path $csvNameMailTipOnly | ForEach-Object { 
    # Set-Mailbox: https://docs.microsoft.com/en-us/powershell/module/exchange/set-mailbox?view=exchange-ps
    $Identity = $_.Identity
    
    # Set MailTip
    $MailTip = $_.MailTip
    Set-Mailbox -Identity $Identity -MailTip $MailTip
    
    Start-Sleep -Seconds 2

    # Set MailTipTranslation German
    $MailTipTranslationsDE = "DE:$($_.MailTipTranslationsDE)"
    Set-Mailbox -Identity $Identity -MailTipTranslations @{Add = $MailTipTranslationsDE }
}

#endregion


#region Set-User
<#
Import-Csv -Path $csvName | ForEach-Object { 
    # Set-User: https://docs.microsoft.com/en-us/powershell/module/exchange/set-user?view=exchange-ps
    
    $paramSetUser = @{
        Identity        = $_.Identity
        Company         = $_.Company
        StreetAddress   = $_.Street
        City            = $_.City
        StateOrProvince = $_.State
        PostalCode      = $_.PostalCode
        CountryOrRegion = $_.CountryOrRegion
        Phone           = $_.Phone
        GeoCoordinates  = $_.GeoCoordinates
    }
    Set-User @paramSetUser
}
#>
#endregion


#region Set-Place
Import-Csv -Path $csvName | ForEach-Object { 
    # Set-Place: https://docs.microsoft.com/en-us/powershell/module/exchange/set-place?view=exchange-ps
    <#
    Note: 
    In hybrid environments, this cmdlet doesn't work on the following properties on synchronized room mailboxes: 
    City, CountryOrRegion, GeoCoordinates, Phone, PostalCode, State, and Street. 
    To modify these properties on synchronized room mailboxes, use the [Set-User] or [Set-Mailbox] cmdlets in on-premises Exchange.
    #>
    $IsWheelChairAccessible = [System.Convert]::ToBoolean($_.IsWheelChairAccessible) 
    $IsMTREnabled = [System.Convert]::ToBoolean($_.MTREnabled)

    $paramSetPlace = @{
        Building               = $_.Building
        Label                  = $_.Label
        Capacity               = $_.Capacity
        AudioDeviceName        = $_.AudioDeviceName
        VideoDeviceName        = $_.VideoDeviceName
        DisplayDeviceName      = $_.DisplayDeviceName
        IsWheelChairAccessible = $IsWheelChairAccessible
        Floor                  = $_.Floor
        FloorLabel             = $_.FloorLabel
        MTREnabled             = $IsMTREnabled #room configured with a Microsoft Teams room system
        Phone                  = $_.Phone
        Tags                   = $_.RoomCategory
        Street                 = $_.Street
        City                   = $_.City
        PostalCode             = $_.PostalCode
        State                  = $_.State
        CountryOrRegion        = $_.CountryOrRegion # https://www.nationsonline.org/oneworld/country_code_list.htm (Alpha 2)
        GeoCoordinates         = $_.GeoCoordinates
    }
    Set-Place -Identity $_.Identity @paramSetPlace
}
#endregion


#region Set-CalendarProcessing
Import-Csv -Path $csvName | ForEach-Object { 
    #Set the Access Permissions for the Service Account
    #Add-MailboxPermission -Identity $_.Identity -User $ServiceAccount -AccessRights FullAccess
    
    # Set-CalendarProcessing: https://docs.microsoft.com/en-us/powershell/module/exchange/set-calendarprocessing?view=exchange-ps
    $AllowRecurringMeetings = [System.Convert]::ToBoolean($_.AllowRecurringMeetings)
    $AllowConflicts = [System.Convert]::ToBoolean($_.AllowConflicts)
    $EnforceSchedulingHorizon = [System.Convert]::ToBoolean($_.EnforceSchedulingHorizon)
    $EnforceCapacity = [System.Convert]::ToBoolean($_.EnforceCapacity)
    $OrganizerInfo = [System.Convert]::ToBoolean($_.OrganizerInfo)
    $ProcessExternalMeetingMessages = [System.Convert]::ToBoolean($_.ProcessExternalMeetingMessages)
    
    $paramSetCalendarProcessing = @{
        # specifies whether to allow recurring meetings in meeting requests
        AllowRecurringMeetings         = $AllowRecurringMeetings
        # specifies whether to allow conflicting meeting requests
        AllowConflicts                 = $AllowConflicts
        # maximum number of conflicts for new recurring meeting requests when [-AllowRecurringMeetings] set to $true
        MaximumConflictInstances       = $_.MaximumConflictInstances
        # maximum percentage of meeting conflicts for new recurring meeting requests
        ConflictPercentageAllowed      = $_.ConflictPercentageAllowed
        # enables or disables calendar processing on the mailbox
        AutomateProcessing             = $_.AutomateProcessing
        # specifies how reservations work on the resource mailbox
        BookingType                    = $_.BookingType
        # maximum number of days in advance that the resource can be reserved
        BookingWindowInDays            = $_.BookingWindowInDays
        # controls the behavior of recurring meetings that extend beyond the date specified by [-BookingWindowInDays]
        EnforceSchedulingHorizon       = $EnforceSchedulingHorizon
        # specifies whether to restrict the number of attendees to the capacity of the Workspace
        EnforceCapacity                = $EnforceCapacity
        # specifies whether the resource mailbox sends organizer information when a meeting request is declined because of conflicts
        OrganizerInfo                  = $OrganizerInfo
        # specifies whether to process meeting requests that originate outside the Exchange organization
        ProcessExternalMeetingMessages = $ProcessExternalMeetingMessages
        # minimum duration in minutes for meeting requests in workspace mailboxes
        MinimumDurationInMinutes       = $_.MinimumDurationInMinutes
        # maximum duration in minutes for meeting requests
        MaximumDurationInMinutes       = $_.MaximumDurationInMinutes
    }
    Set-CalendarProcessing -Identity $_.Identity @paramSetCalendarProcessing
}
#endregion


#region Set-CalendarProcessing - AdditionalResponse (HTML)
# Set-CalendarProcessing -AdditionalResponse
# Set-CalendarProcessing: https://docs.microsoft.com/en-us/powershell/module/exchange/set-calendarprocessing?view=exchange-ps
[String]$AdditionalResponse = Get-Content -Path .\EXUSG_Set-CalendarProcessing_AdditionalResponse.html
Import-Csv -Path $csvName | ForEach-Object { 
    
    $Identity = $_.Identity
    
    switch ($_.Type) {
        'Room' { 
            # Set AdditionalResponse as HTML file
            Write-Host "Room $Identity - AdditionalResponse set"
            $paramSetMailboxAdditionalResponse = @{
                AddAdditionalResponse = $true
                AdditionalResponse    = $AdditionalResponse
            };
            Set-CalendarProcessing -Identity $Identity @paramSetMailboxAdditionalResponse

        }
        'Workspace' { 
            Write-Host "Workspace $Identity - AdditionalResponse not required" -ForegroundColor Green;
            Set-CalendarProcessing -Identity $Identity -AddAdditionalResponse $false
        }
        default { 
            Write-Error 'Either there is no Mailbox Type attribute (Room or Workspace) maintained or there is an error'
        }
    }
}
#endregion


#region New-DistributionGroup - Roomlist
New-DistributionGroup -Name 'DL.MR.DE.LOC.BuildingA'-DisplayName 'Building A' -RoomList
New-DistributionGroup -Name 'DL.MR.DE.LOC.BuildingB'-DisplayName 'Building B' -RoomList
New-DistributionGroup -Name 'DL.MR.DE.LOC.BuildingC'-DisplayName 'Building C' -RoomList
#endregion

 
#region Add-DistributionGroupMember
Import-Csv -Path $csvName | ForEach-Object { 
    # Set-User: https://docs.microsoft.com/en-us/powershell/module/exchange/set-user?view=exchange-ps
    $Identity = $_.Identity
    switch ($_.Building) {
        'A' { Add-DistributionGroupMember -Identity 'DL.MR.DE.LOC.BuildingA' -Member $Identity }
        'B' { Add-DistributionGroupMember -Identity 'DL.MR.DE.LOC.BuildingB' -Member $Identity }
        'C' { Add-DistributionGroupMember -Identity 'DL.MR.DE.LOC.BuildingC' -Member $Identity }
        default { Write-Error 'Either there is no building maintained or there is an error' }
    }
}

# Get all RoomLists
Get-DistributionGroup -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -eq 'RoomList' } | Format-Table DisplayName,Identity,PrimarySmtpAddress -AutoSize

Get-DistributionGroupMember -Identity 'DL.MR.DE.LOC.BuildingB'
#endregion


#region Update-DistributionGroupMember -Identity 'DL.MR.DE.LOC.BuildingA' -Members $MembersA -Confirm:$false
# [string[]]$MembersA = 'MR.DE.LOC.017@dedas.onmicrosoft.com','MR.DE.LOC.018@dedas.onmicrosoft.com'
# Update-DistributionGroupMember -Identity 'DL.MR.DE.LOC.BuildingA' -Members $MembersA -Confirm:$false
#endregion


#region Restrict / Block who can Book a Room
# Add old Rooms rooms to string array
[string[]]$oldRooms = 'MR.DE.LOC.017@dedas.onmicrosoft.com','MR.DE.LOC.018@dedas.onmicrosoft.com'

# Add users allowed to book old Rooms to string array
[string[]]$BookInPolicy = 'admin@dedas.onmicrosoft.com','AlexW@dedas.onmicrosoft.com'

$oldRooms | ForEach-Object { 

    $oldRoomMailbox = Get-EXOMailbox -Identity $_
    # Set-CalendarProcessing: https://docs.microsoft.com/en-us/powershell/module/exchange/set-calendarprocessing?view=exchange-ps
    $paramSetCalendarProcessingOldRoom = @{
        AllBookInPolicy    = $false
        AllRequestInPolicy = $false 
        BookInPolicy       = $BookInPolicy
        ResourceDelegates  = $BookInPolicy
    }
    Set-CalendarProcessing -Identity $oldRoomMailbox.Identity @paramSetCalendarProcessingOldRoom
    
    # Set-Mailbox: https://docs.microsoft.com/en-us/powershell/module/exchange/set-mailbox?view=exchange-ps
    $paramSetMailboxOldRoom = @{
        MailTip     = 'Dieser Raum wird im Zuge des Umzugs entfernt. Weitere Infos bei IT erfragen.'
        DisplayName = "$($oldRoomMailbox.DisplayName) [NICHT VERWENDEN]"
    }
    Set-Mailbox -Identity $oldRoomMailbox.Identity @paramSetMailboxOldRoom

    # Add FullAccess to Mailbox for ResourceDelegates
    $BookInPolicy | ForEach-Object {
        $DelegateUser = Get-EXOMailbox -Identity $_
        #Set the Access Permissions for the Service Account
        Add-MailboxPermission -Identity $oldRoomMailbox.Identity -User $DelegateUser -AccessRights FullAccess -WarningAction SilentlyContinue
    }

    Get-CalendarProcessing -Identity $oldRoomMailbox | Select-Object Identity,ResourceDelegates,BookInPolicy
}
#endregion


#region MS GE - List places
# Microsoft Graph Explorer - aka.ms/ge
# https://learn.microsoft.com/en-us/graph/api/place-list?view=graph-rest-1.0&tabs=http

# Permissions: Place.Read.All
# https://learn.microsoft.com/en-us/graph/api/place-list?view=graph-rest-1.0&tabs=http#permissions

# Get all the rooms in a tenant (no Workspaces!)
GET https://graph.microsoft.com/v1.0/places/microsoft.graph.room

# Get all the workspaces in a tenant (beta only)
GET https://graph.microsoft.com/beta/places/microsoft.graph.workspace

# Get all the room lists in a tenant
GET https://graph.microsoft.com/v1.0/places/microsoft.graph.roomlist

#endregion


#region MS GE - Get place - Get a room
# Microsoft Graph Explorer - aka.ms/ge
# https://learn.microsoft.com/en-us/graph/api/place-get?view=graph-rest-1.0&tabs=http

# Permissions: Place.Read.All
# https://learn.microsoft.com/en-us/graph/api/place-list?view=graph-rest-1.0&tabs=http#permissions

# Get a room by place-id
GET https://graph.microsoft.com/v1.0/places/1f63500f-5c67-491a-b198-699b1ce5debc
GET https://graph.microsoft.com/v1.0/places/MR.DE.LOC.002@dedas.onmicrosoft.com

# Get all the room lists in a tenant
GET https://graph.microsoft.com/v1.0/places/DL.MR.DE.LOC.BuildingA@dedas.onmicrosoft.com

#endregion


#region MS GE - Get place - Update place
# Microsoft Graph Explorer - aka.ms/ge
# https://learn.microsoft.com/en-us/graph/api/place-update?view=graph-rest-1.0&tabs=http

# Permissions: Place.ReadWrite.All
# https://learn.microsoft.com/en-us/graph/api/place-update?view=graph-rest-1.0&tabs=http#permissions

# Update a room by place-id
PATCH https://graph.microsoft.com/v1.0/places/MR.DE.LOC.004@dedas.onmicrosoft.com

<#
Content-type: application/json

{
  "@odata.type": "microsoft.graph.room",
  "phone": "+49 30 999 99 000",
  "capacity": 10,
  "isWheelChairAccessible": true,
  "tags": [
        "Meeting Room",
        "Block seating"
    ],
    "audioDeviceName": "Jabra Teams Room System",
    "videoDeviceName": "Jabra Teams Room System",
    "displayDeviceName": "Samsung Smart TV 60 inch"
}
#>

# Get all the room lists in a tenant
GET https://graph.microsoft.com/v1.0/places/DL.MR.DE.LOC.BuildingA@dedas.onmicrosoft.com


#endregion