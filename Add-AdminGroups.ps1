# Bootstrap local Administrators group with central service teams.
# Update the list below to match your directory groups.
$supportGroups = @(
    'CONTOSO\EndpointWorkstationAdmins',
    'CONTOSO\ServiceDeskOperators',
    'CONTOSO\FieldITSpecialists',
    'CONTOSO\OnsiteSupport',
    'CONTOSO\TemporaryAdmins'
)

foreach ($group in $supportGroups) {
    try {
        Add-LocalGroupMember -Group 'Administrators' -Member $group -ErrorAction Stop
        Write-Host "Added $group to local Administrators."
    } catch {
        Write-Warning "Could not add $group: $($_.Exception.Message)"
    }
}
