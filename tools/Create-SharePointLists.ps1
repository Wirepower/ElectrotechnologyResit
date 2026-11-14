# Creates the Electrotechnology Resit SharePoint lists if they are missing.
# Safe to run again. This script never deletes a list, a column, or a row.
# The app's SharePoint setup wizard is the supported path. Use this only if IT needs PnP PowerShell.
# Requires: Install-Module PnP.PowerShell
# Example:
#   ./Create-SharePointLists.ps1 -SiteUrl "https://tenant.sharepoint.com/sites/YourSite"

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $SiteUrl,
    [string] $ResitBookingsList = "ResitBookings",
    [string] $StudentsList = "Students",
    [string] $TeachersList = "Teachers",
    [string] $UnitsList = "Units",
    [string] $SignatureLibrary = "EmailSignature"
)

$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -Interactive

function Ensure-ResitList {
    param([string] $Title, [string] $Template)
    $existing = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if (-not $existing) {
        New-PnPList -Title $Title -Template $Template | Out-Null
        Write-Host "Created $Template '$Title'."
    } else {
        Write-Host "Left existing list '$Title' in place."
    }
}

function Ensure-TextColumn {
    param([string] $List, [string] $Internal, [string] $Display, [switch] $Indexed, [switch] $Unique)
    $field = Get-PnPField -List $List -Identity $Internal -ErrorAction SilentlyContinue
    if (-not $field) {
        Add-PnPField -List $List -DisplayName $Display -InternalName $Internal -Type Text -AddToDefaultView | Out-Null
        $field = Get-PnPField -List $List -Identity $Internal
        Write-Host "Added $List.$Internal."
    }
    $changed = $false
    if ($Indexed -and -not $field.Indexed) { $field.Indexed = $true; $changed = $true }
    if ($Unique -and -not $field.EnforceUniqueValues) {
        $field.Indexed = $true
        $field.EnforceUniqueValues = $true
        $changed = $true
    }
    if ($changed) {
        $field.Update()
        Invoke-PnPQuery
        Write-Host "Updated index or unique setting on $List.$Internal."
    }
}

function Ensure-DateColumn {
    param([string] $List, [string] $Internal, [string] $Display)
    $field = Get-PnPField -List $List -Identity $Internal -ErrorAction SilentlyContinue
    if (-not $field) {
        Add-PnPField -List $List -DisplayName $Display -InternalName $Internal -Type DateTime -AddToDefaultView | Out-Null
        $field = Get-PnPField -List $List -Identity $Internal
        $schema = $field.SchemaXml
        $schema = $schema -replace 'Format="DateTime"', 'Format="DateOnly"'
        $field.SchemaXml = $schema
        $field.Update()
        Invoke-PnPQuery
        Write-Host "Added $List.$Internal."
        $field = Get-PnPField -List $List -Identity $Internal
    }
    if (-not $field.Indexed) {
        $field.Indexed = $true
        $field.Update()
        Invoke-PnPQuery
        Write-Host "Indexed $List.$Internal."
    }
}

function Ensure-YesNoColumn {
    param([string] $List, [string] $Internal, [string] $Display)
    $field = Get-PnPField -List $List -Identity $Internal -ErrorAction SilentlyContinue
    if (-not $field) {
        Add-PnPField -List $List -DisplayName $Display -InternalName $Internal -Type Boolean -AddToDefaultView | Out-Null
        Write-Host "Added $List.$Internal."
    }
}

Ensure-ResitList -Title $ResitBookingsList -Template GenericList
Ensure-TextColumn -List $ResitBookingsList -Internal "StudentID" -Display "Student ID" -Unique
Ensure-TextColumn -List $ResitBookingsList -Internal "StudentFirstname" -Display "Student Firstname"
Ensure-TextColumn -List $ResitBookingsList -Internal "StudentSurname" -Display "Student Surname"
Ensure-TextColumn -List $ResitBookingsList -Internal "StudentEmail" -Display "Student Email"
Ensure-TextColumn -List $ResitBookingsList -Internal "Unit" -Display "Unit"
Ensure-TextColumn -List $ResitBookingsList -Internal "UnitName" -Display "Unit Name"
Ensure-TextColumn -List $ResitBookingsList -Internal "Blockgroup" -Display "Blockgroup"
Ensure-TextColumn -List $ResitBookingsList -Internal "AllocatedTeacher" -Display "Allocated Teacher"
Ensure-TextColumn -List $ResitBookingsList -Internal "AllocatedTeacherEmail" -Display "Allocated Teacher Email"
Ensure-TextColumn -List $ResitBookingsList -Internal "AttemptNo" -Display "Attempt No"
Ensure-DateColumn -List $ResitBookingsList -Internal "ResitDate" -Display "Resit date"
Ensure-YesNoColumn -List $ResitBookingsList -Internal "EnergyspaceCreated" -Display "Energyspace Created"
Ensure-YesNoColumn -List $ResitBookingsList -Internal "EnergyspaceAssessmentBooked" -Display "Energyspace Assessment Booked"
Ensure-TextColumn -List $ResitBookingsList -Internal "Status" -Display "status"

Ensure-ResitList -Title $StudentsList -Template GenericList
Ensure-TextColumn -List $StudentsList -Internal "StudentID" -Display "Student ID" -Unique
Ensure-TextColumn -List $StudentsList -Internal "StudentGivenName" -Display "Student Given Name"
Ensure-TextColumn -List $StudentsList -Internal "StudentFamilyName" -Display "Student Family Name"
Ensure-TextColumn -List $StudentsList -Internal "StudentPersonalEmail" -Display "Student Personal Email"
Ensure-TextColumn -List $StudentsList -Internal "BlockGroupCode" -Display "Block Group Code"

Ensure-ResitList -Title $TeachersList -Template GenericList
Ensure-TextColumn -List $TeachersList -Internal "TeacherFullName" -Display "Teacher Full Name" -Indexed
Ensure-TextColumn -List $TeachersList -Internal "Email" -Display "Email"

Ensure-ResitList -Title $UnitsList -Template GenericList
Ensure-TextColumn -List $UnitsList -Internal "UnitCode" -Display "Unit Code" -Unique
Ensure-TextColumn -List $UnitsList -Internal "UnitTitle" -Display "Unit Title"

Ensure-ResitList -Title $SignatureLibrary -Template DocumentLibrary

Write-Host "Finished. Existing data and columns were not deleted."
