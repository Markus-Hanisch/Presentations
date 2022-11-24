#region Connect to Microsoft Graph with required Scopes

# Import-Module Microsoft.Graph.Calendar


$RequiredScopes = @('Place.Read.All','Place.ReadWrite.All','User.Read')
Connect-MgGraph -Scopes $RequiredScopes

Select-MgProfile beta # in production use v1
$Details = Get-MgContext
$Scopes = $Details | Select-Object -ExpandProperty Scopes
$Scopes = $Scopes -Join ', '
$ProfileName = (Get-MgProfile).Name

Write-Host 'Microsoft Graph Connection Information'
Write-Host '--------------------------------------'
Write-Host ' '
Write-Host '+-------------------------------------------------------------------------------------------------------------------+'
Write-Host ('Profile set as {0}. The following permission scope is defined: {1}' -f $ProfileName, $Scopes)
Write-Host ''

#endregion


$MgPlaceId = 'a87fe657-3abc-46a0-9b91-0cd71bf600fb' # MR.DE.LOC.016@dedas.onmicrosoft.com

#region Get-MgPlace for specific PlaceId
Get-MgPlace -PlaceId $MgPlaceId
#endregion


#region Update-MgPlace for specific PlaceId
$params = @{
    '@odata.type' = 'microsoft.graph.room'
    capacity = 10
    tags                   = @(
        'Focused Work',
        'U shape seating'
    )
}

Update-MgPlace -PlaceId $MgPlaceId -BodyParameter $params
#endregion

Disconnect-MgGraph | Out-Null