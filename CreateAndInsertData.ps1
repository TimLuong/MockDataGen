# PowerShell script to import sample data into the Medical System SharePoint lists
# This script generates realistic medical data for testing and demonstration
# Order: Patients → Doctors → Appointments → PatientJourneyActivities
# Updated for Power Apps SharePoint lookup field compatibility
# FIXED VERSION - Corrected lookup field references and error handling
# ENHANCED VERSION - Auto-creates lists if they don't exist with proper structure

param(
    [string]$SiteUrl = "https://mngenvmcap432670.sharepoint.com/sites/ProjectManagement",
    [switch]$ClearExistingData,
    [switch]$CreateListsIfMissing
)

# Set default value for CreateListsIfMissing if not specified
if (-not $PSBoundParameters.ContainsKey('CreateListsIfMissing')) {
    $CreateListsIfMissing = $true
}

function Test-LastCommand {
    param([string]$Action)
    if ($?) {
        Write-Host "$Action - Success" -ForegroundColor Green
    } else {
        Write-Host "$Action - Failed: $($Error[0])" -ForegroundColor Red
        exit 1
    }
}

function New-PatientsListIfMissing {
    Write-Host "Creating Patients list..." -ForegroundColor Cyan
   
    # Remove existing list if it exists to ensure clean slate
    $existingList = Get-PnPList -Identity "Patients" -ErrorAction SilentlyContinue
    if ($existingList) {
        Write-Host "Removing existing 'Patients' list to recreate with correct structure..." -ForegroundColor Yellow
        Remove-PnPList -Identity "Patients" -Force
    }
   
    New-PnPList -Title "Patients" -Template GenericList
    Test-LastCommand "Create 'Patients' list"

    # Configure Title field for display only (not for ID)
    Set-PnPField -List "Patients" -Identity "Title" -Values @{Title="DisplayName"; Required=$false}

    # Add PatientID field (custom formatted, unique, indexed)
    Add-PnPField -List "Patients" -DisplayName "PatientID" -InternalName "PatientID" -Type Text -AddToDefaultView -Required
    Set-PnPField -List "Patients" -Identity "PatientID" -Values @{EnforceUniqueValues=$true; Indexed=$true}

    # === 1. Add all standard (non-calculated, non-lookup) fields ===
    Add-PnPField -List "Patients" -DisplayName "FirstName" -InternalName "FirstName" -Type Text -AddToDefaultView
    Add-PnPField -List "Patients" -DisplayName "LastName" -InternalName "LastName" -Type Text -AddToDefaultView
    Add-PnPField -List "Patients" -DisplayName "DateOfBirth" -InternalName "DateOfBirth" -Type Text -AddToDefaultView
    Add-PnPField -List "Patients" -DisplayName "Gender" -InternalName "Gender" -Type Choice -Choices @("Male","Female","Other") -AddToDefaultView
    Add-PnPField -List "Patients" -DisplayName "ContactNumber" -InternalName "ContactNumber" -Type Text
    Add-PnPField -List "Patients" -DisplayName "EmailAddress" -InternalName "EmailAddress" -Type Text
    Add-PnPField -List "Patients" -DisplayName "Address" -InternalName "Address" -Type Text
    Add-PnPField -List "Patients" -DisplayName "MedicalHistorySummary" -InternalName "MedicalHistorySummary" -Type Note
    Add-PnPField -List "Patients" -DisplayName "PatientStatus" -InternalName "PatientStatus" -Type Choice -Choices @("New","In Treatment","Awaiting Follow-up","High Priority","Discharged") -AddToDefaultView

    # === 2. Add calculated fields (after all referenced fields exist) ===
    Add-PnPField -List "Patients" -DisplayName "FullName" -InternalName "FullName" -Type Calculated -Formula '=[FirstName] &amp; " " &amp; [LastName]' -AddToDefaultView

    Write-Host "✅ Patients list created successfully!" -ForegroundColor Green
}

function New-DoctorsListIfMissing {
    Write-Host "Creating Doctors list..." -ForegroundColor Cyan
   
    # Remove existing list if it exists to ensure clean slate
    $existingList = Get-PnPList -Identity "Doctors" -ErrorAction SilentlyContinue
    if ($existingList) {
        Write-Host "Removing existing 'Doctors' list to recreate with correct structure..." -ForegroundColor Yellow
        Remove-PnPList -Identity "Doctors" -Force
    }
   
    New-PnPList -Title "Doctors" -Template GenericList
    Test-LastCommand "Create 'Doctors' list"

    # Configure Title field for display only (not for ID)
    Set-PnPField -List "Doctors" -Identity "Title" -Values @{Title="DisplayName"; Required=$false}

    # Add DoctorID field (custom formatted, unique, indexed)
    Add-PnPField -List "Doctors" -DisplayName "DoctorID" -InternalName "DoctorID" -Type Text -AddToDefaultView -Required
    Set-PnPField -List "Doctors" -Identity "DoctorID" -Values @{EnforceUniqueValues=$true; Indexed=$true}

    # Add other required fields
    Add-PnPField -List "Doctors" -DisplayName "FirstName" -InternalName "FirstName" -Type Text -AddToDefaultView
    Add-PnPField -List "Doctors" -DisplayName "LastName" -InternalName "LastName" -Type Text -AddToDefaultView
    Add-PnPField -List "Doctors" -DisplayName "Specialization" -InternalName "Specialization" -Type Text -AddToDefaultView
    Add-PnPField -List "Doctors" -DisplayName "ContactEmail" -InternalName "ContactEmail" -Type Text
    Add-PnPField -List "Doctors" -DisplayName "Department" -InternalName "Department" -Type Text -AddToDefaultView

    # Calculated field for FullName
    Add-PnPField -List "Doctors" -DisplayName "FullName" -InternalName "FullName" -Type Calculated -Formula '=[FirstName] &amp; " " &amp; [LastName]' -AddToDefaultView

    Write-Host "✅ Doctors list created successfully!" -ForegroundColor Green
}

function New-AppointmentsListIfMissing {
    Write-Host "Creating Appointments list..." -ForegroundColor Cyan
   
    # Remove existing list if it exists to ensure clean slate
    $existingList = Get-PnPList -Identity "Appointments" -ErrorAction SilentlyContinue
    if ($existingList) {
        Write-Host "Removing existing 'Appointments' list to recreate with correct structure..." -ForegroundColor Yellow
        Remove-PnPList -Identity "Appointments" -Force
    }
   
    New-PnPList -Title "Appointments" -Template GenericList
    Test-LastCommand "Create 'Appointments' list"

    # Configure Title field for display only (not for ID)
    Set-PnPField -List "Appointments" -Identity "Title" -Values @{Title="DisplayName"; Required=$false}

    # Add AppointmentID field (custom formatted, unique, indexed)
    Add-PnPField -List "Appointments" -DisplayName "AppointmentID" -InternalName "AppointmentID" -Type Text -AddToDefaultView -Required
    Set-PnPField -List "Appointments" -Identity "AppointmentID" -Values @{EnforceUniqueValues=$true; Indexed=$true}

    # === 1. Add all standard (non-calculated, non-lookup) fields ===
    Add-PnPField -List "Appointments" -DisplayName "AppointmentDateTime" -InternalName "AppointmentDateTime" -Type DateTime -AddToDefaultView -Required
    Add-PnPField -List "Appointments" -DisplayName "AppointmentEndTime" -InternalName "AppointmentEndTime" -Type DateTime -AddToDefaultView
    Add-PnPField -List "Appointments" -DisplayName "ServiceType" -InternalName "ServiceType" -Type Choice -Choices @("Consultation","Follow-up","Lab Test","Imaging","Surgery","Therapy","Vaccination","Screening","Physical Exam","Discharge") -AddToDefaultView
    Add-PnPField -List "Appointments" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices @("Scheduled","Confirmed","Rescheduled","Completed","No Show","Cancelled") -AddToDefaultView
    Add-PnPField -List "Appointments" -DisplayName "Notes" -InternalName "Notes" -Type Note
    Add-PnPField -List "Appointments" -DisplayName "IsUrgent" -InternalName "IsUrgent" -Type Boolean

    # === 2. Add lookup fields that reference calculated FullName fields ===
    Add-PnPField -List "Appointments" -DisplayName "PatientFullName" -InternalName "PatientFullNameLookup" -Type Lookup -AddToDefaultView -Required
    Write-Host "Configuring PatientFullName lookup field..." -ForegroundColor Yellow
    $patientsList = Get-PnPList -Identity "Patients"
    Set-PnPField -List "Appointments" -Identity "PatientFullNameLookup" -Values @{LookupList=$patientsList.Id.ToString(); LookupField="FullName"; Indexed=$true}
    Test-LastCommand "Configure PatientFullName lookup field"

    Add-PnPField -List "Appointments" -DisplayName "DoctorFullName" -InternalName "DoctorFullNameLookup" -Type Lookup -AddToDefaultView -Required
    Write-Host "Configuring DoctorFullName lookup field..." -ForegroundColor Yellow
    $doctorsList = Get-PnPList -Identity "Doctors"
    Set-PnPField -List "Appointments" -Identity "DoctorFullNameLookup" -Values @{LookupList=$doctorsList.Id.ToString(); LookupField="FullName"; Indexed=$true}
    Test-LastCommand "Configure DoctorFullName lookup field"

    Write-Host "✅ Appointments list created successfully!" -ForegroundColor Green
}

function New-PatientJourneyActivitiesListIfMissing {
    Write-Host "Creating PatientJourneyActivities list..." -ForegroundColor Cyan
   
    # Remove existing list if it exists to ensure clean slate
    $existingList = Get-PnPList -Identity "PatientJourneyActivities" -ErrorAction SilentlyContinue
    if ($existingList) {
        Write-Host "Removing existing 'PatientJourneyActivities' list to recreate with correct structure..." -ForegroundColor Yellow
        Remove-PnPList -Identity "PatientJourneyActivities" -Force
    }
   
    New-PnPList -Title "PatientJourneyActivities" -Template GenericList
    Test-LastCommand "Create 'PatientJourneyActivities' list"

    # Configure Title field for display only (not for ID)
    Set-PnPField -List "PatientJourneyActivities" -Identity "Title" -Values @{Title="DisplayName"; Required=$false}

    # Add ActivityID field (custom formatted, unique, indexed)
    Add-PnPField -List "PatientJourneyActivities" -DisplayName "ActivityID" -InternalName "ActivityID" -Type Text -AddToDefaultView -Required
    Set-PnPField -List "PatientJourneyActivities" -Identity "ActivityID" -Values @{EnforceUniqueValues=$true; Indexed=$true}

    # === 1. Add all standard (non-calculated, non-lookup) fields ===
    Add-PnPField -List "PatientJourneyActivities" -DisplayName "ActivityDateTime" -InternalName "ActivityDateTime" -Type DateTime -AddToDefaultView -Required
    Add-PnPField -List "PatientJourneyActivities" -DisplayName "ActivityType" -InternalName "ActivityType" -Type Choice -Choices @("Admission","Consultation","Lab Test","Medication","Treatment","Imaging","Follow-up Call","Discharge") -AddToDefaultView
    Add-PnPField -List "PatientJourneyActivities" -DisplayName "Notes" -InternalName "Notes" -Type Note
    Add-PnPField -List "PatientJourneyActivities" -DisplayName "Duration" -InternalName "Duration" -Type Number
    Add-PnPField -List "PatientJourneyActivities" -DisplayName "Priority" -InternalName "Priority" -Type Choice -Choices @("Normal","High","Low","Urgent") -AddToDefaultView

    # === 2. Add calculated fields (after all referenced fields exist) ===
    Add-PnPField -List "PatientJourneyActivities" -DisplayName "ActivityTitle" -InternalName "ActivityTitle" -Type Calculated -Formula '=[ActivityType] &amp; " - " &amp; [Priority]' -AddToDefaultView

    # === 3. Add lookup fields (after all referenced fields exist) ===
    Add-PnPField -List "PatientJourneyActivities" -DisplayName "PatientFullName" -InternalName "PatientFullNameLookup" -Type Lookup -AddToDefaultView -Required
    Write-Host "Configuring PatientFullName lookup field..." -ForegroundColor Yellow
    $patientsList = Get-PnPList -Identity "Patients"
    Set-PnPField -List "PatientJourneyActivities" -Identity "PatientFullNameLookup" -Values @{LookupList=$patientsList.Id.ToString(); LookupField="FullName"; Indexed=$true}
    Test-LastCommand "Configure PatientFullName lookup field"

    Add-PnPField -List "PatientJourneyActivities" -DisplayName "DoctorFullName" -InternalName "DoctorFullNameLookup" -Type Lookup -AddToDefaultView
    Write-Host "Configuring DoctorFullName lookup field..." -ForegroundColor Yellow
    $doctorsList = Get-PnPList -Identity "Doctors"
    Set-PnPField -List "PatientJourneyActivities" -Identity "DoctorFullNameLookup" -Values @{LookupList=$doctorsList.Id.ToString(); LookupField="FullName"; Indexed=$true}
    Test-LastCommand "Configure DoctorFullName lookup field"

    Add-PnPField -List "PatientJourneyActivities" -DisplayName "AppointmentReference" -InternalName "AppointmentIDLookup" -Type Lookup -AddToDefaultView
    Write-Host "Configuring AppointmentReference lookup field..." -ForegroundColor Yellow
    $appointmentsList = Get-PnPList -Identity "Appointments"
    Set-PnPField -List "PatientJourneyActivities" -Identity "AppointmentIDLookup" -Values @{LookupList=$appointmentsList.Id.ToString(); LookupField="AppointmentID"; Indexed=$true}
    Test-LastCommand "Configure AppointmentReference lookup field"

    Write-Host "✅ PatientJourneyActivities list created successfully!" -ForegroundColor Green
}

function Initialize-SharePointListsIfNeeded {
    if ($CreateListsIfMissing) {
        Write-Host "🔍 Checking and creating SharePoint lists if missing..." -ForegroundColor Cyan
       
        try {
            New-PatientsListIfMissing
            New-DoctorsListIfMissing
            New-AppointmentsListIfMissing
            New-PatientJourneyActivitiesListIfMissing
           
            Write-Host "✅ All required SharePoint lists are now available!" -ForegroundColor Green
        }
        catch {
            Write-host "❌ Error creating lists: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "You may need to run the PowerApps-Compatible-Lists.ps1 script manually first." -ForegroundColor Yellow
            throw
        }
    }
}

function Get-RandomDate {
    param(
        [DateTime]$StartDate,
        [DateTime]$EndDate
    )
    $timeSpan = $EndDate - $StartDate
    $randomDays = Get-Random -Minimum 0 -Maximum $timeSpan.Days
    $randomHours = Get-Random -Minimum 8 -Maximum 17  # Business hours 8 AM to 5 PM
    $randomMinutes = @(0, 15, 30, 45) | Get-Random   # Quarter-hour appointments
    return $StartDate.AddDays($randomDays).AddHours($randomHours).AddMinutes($randomMinutes)
}

function Clear-ListData {
    param([string]$ListName)
   
    Write-Host "Clearing existing data from $ListName..." -ForegroundColor Yellow
    try {
        $items = Get-PnPListItem -List $ListName
        if ($items.Count -gt 0) {
            foreach ($item in $items) {
                Remove-PnPListItem -List $ListName -Identity $item.Id -Force
            }
            Write-Host "  Cleared $($items.Count) items from $ListName" -ForegroundColor Green
        } else {
            Write-Host "  $ListName is already empty" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "  Warning: Could not clear $ListName - $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Connect to SharePoint with ClientId
Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId 895298f8-3468-445f-8059-0d925de101fc
Test-LastCommand "Connect to SharePoint"

# Ensure all required SharePoint lists exist
Initialize-SharePointListsIfNeeded

Write-Host "Starting expanded data import for Medical System..." -ForegroundColor Cyan

# === CLEAR EXISTING DATA IF REQUESTED ===
if ($ClearExistingData) {
    Write-Host "Clearing existing data from all lists..." -ForegroundColor Yellow
    Clear-ListData "PatientJourneyActivities"
    Clear-ListData "Appointments"
    Clear-ListData "Doctors"
    Clear-ListData "Patients"
    Write-Host "Data clearing complete." -ForegroundColor Green
}

# === ID Format Generators ===
# Utility functions for ID generation
function Get-PatientID {
    param([int]$n)
    return ('MRN{0:D5}' -f $n)
}
function Get-DoctorID {
    param([int]$n)
    return ('DOC{0:D4}' -f $n)
}
function Get-AppointmentID {
    param([int]$n)
    return ('APP{0:D6}' -f $n)
}
function Get-ActivityID {
    param([int]$n)
    return ('ACT{0:D6}' -f $n)
}

# === PATIENTS DATA GENERATION ===
$firstNames = @("John", "Jane", "Michael", "Emily", "David", "Sarah", "Chris", "Jessica", "Daniel", "Ashley", "Matthew", "Amanda", "Joshua", "Brittany", "Andrew", "Samantha", "Joseph", "Lauren", "Nicholas", "Megan")
$lastNames = @("Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin")
$genders = @("Male", "Female", "Other")
$medicalConditions = @("Hypertension", "Diabetes", "Asthma", "COPD", "Heart Disease", "Arthritis", "Depression", "Anxiety", "Obesity", "Cancer")
$patientStatuses = @("New", "In Treatment", "Awaiting Follow-up", "High Priority", "Discharged")

$patients = @()
for ($i = 1; $i -le 30; $i++) {
    $firstName = $firstNames | Get-Random
    $lastName = $lastNames | Get-Random
    $fullName = "$firstName $lastName"  # Create full name for Title field
   
    $patients += @{
        PatientID = Get-PatientID $i
        FirstName = $firstName
        LastName = $lastName
        Title = $fullName  # Title field will display the full name for lookups
        DateOfBirth = (Get-RandomDate -StartDate (Get-Date).AddYears(-50) -EndDate (Get-Date).AddYears(-18)).ToString("yyyy-MM-dd")
        Gender = $genders | Get-Random
        ContactNumber = "+1-555-{0:D4}" -f (Get-Random -Minimum 1000 -Maximum 9999)
        EmailAddress = "$($firstName.ToLower()).$($lastName.ToLower())@email.com"
        Address = "{0} {1} St, City, State {2:D5}" -f (Get-Random -Minimum 100 -Maximum 9999), ($lastNames | Get-Random), (Get-Random -Minimum 10000 -Maximum 99999)
        MedicalHistorySummary = $medicalConditions | Get-Random
        PatientStatus = $patientStatuses | Get-Random
    }
}

$patientCount = 0
foreach ($patient in $patients) {
    try {
        Add-PnPListItem -List "Patients" -Values $patient | Out-Null
        $patientCount++
        Write-Host "  Added patient: $($patient.Title)" -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed to add patient: $($patient.Title) - $($_.Exception.Message)" -ForegroundColor Red
    }
}

# === DOCTORS DATA GENERATION ===
$doctorFirstNames = @("James", "Mary", "Robert", "Patricia", "John", "Jennifer", "William", "Linda", "Charles", "Elizabeth", "Joseph", "Barbara", "Thomas", "Susan", "Christopher", "Jessica", "Daniel", "Sarah", "Paul", "Karen")
$specializations = @("Cardiology", "Dermatology", "Endocrinology", "Gastroenterology", "Neurology", "Oncology", "Pediatrics", "Psychiatry", "Radiology", "Surgery")
$departments = @("General Medicine", "Surgery", "Pediatrics", "Radiology", "Oncology", "Cardiology", "Neurology", "Emergency", "Orthopedics", "Dermatology")

$doctors = @()
for ($i = 1; $i -le 10; $i++) {
    $firstName = $doctorFirstNames | Get-Random
    $lastName = $lastNames | Get-Random
    $fullName = "Dr. $firstName $lastName"  # Create full name for Title field with Dr. prefix
   
    $doctors += @{
        DoctorID = Get-DoctorID $i
        FirstName = $firstName
        LastName = $lastName
        Title = $fullName  # Title field will display the full name for lookups
        Specialization = $specializations | Get-Random
        ContactEmail = "$($firstName.Substring(0,1).ToLower()).$($lastName.ToLower())@hospital.com"
        Department = $departments | Get-Random
    }
}

$doctorCount = 0
foreach ($doctor in $doctors) {
    try {
        Add-PnPListItem -List "Doctors" -Values $doctor | Out-Null
        $doctorCount++
        Write-Host "  Added doctor: $($doctor.Title)" -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed to add doctor: $($doctor.Title) - $($_.Exception.Message)" -ForegroundColor Red
    }
}

# === GET LOOKUP IDs FOR APPOINTMENTS ===
Write-Host "Retrieving Patient and Doctor IDs for lookup relationships..." -ForegroundColor Yellow

$patients = Get-PnPListItem -List "Patients"
$doctors = Get-PnPListItem -List "Doctors"

Write-Host "  Found $($patients.Count) patients and $($doctors.Count) doctors" -ForegroundColor Cyan

# Create lookup maps using PatientID and DoctorID
$patientLookup = @{}
$doctorLookup = @{}
foreach ($patient in $patients) {
    if ($patient["PatientID"]) {
        $patientLookup[$patient["PatientID"]] = $patient.Id
    }
}
foreach ($doctor in $doctors) {
    if ($doctor["DoctorID"]) {
        $doctorLookup[$doctor["DoctorID"]] = $doctor.Id
    }
}
Write-Host "  Created lookup maps: $($patientLookup.Count) patients, $($doctorLookup.Count) doctors" -ForegroundColor Cyan

# Define date range for appointments
$startDate = Get-Date "2025-05-15"
$endDate = Get-Date "2025-07-16"

# === APPOINTMENTS DATA GENERATION ===
$serviceTypes = @("Consultation", "Follow-up", "Lab Test", "Imaging", "Surgery", "Therapy", "Vaccination", "Screening", "Physical Exam", "Discharge")

$appointments = @()
for ($i = 1; $i -le 30; $i++) {
    $patient = $patients[($i-1) % $patients.Count]
    $doctor = $doctors[($i-1) % $doctors.Count]
    $date = (Get-Date).AddDays((Get-Random -Minimum 0 -Maximum 30))
    $apptID = Get-AppointmentID $i
    $serviceType = $serviceTypes | Get-Random
   
    # Determine lookup IDs for FullName references
    $patientLookupId = $patientLookup[$patient["PatientID"]]
    $doctorLookupId  = $doctorLookup[$doctor["DoctorID"]]

    # Create appointment title for display using Title field values
    $apptTitle = "$serviceType for $($patient["Title"]) with $($doctor["Title"]) on $($date.ToString('yyyy-MM-dd'))"
   
    $appointments += @{
        Title                    = $apptTitle
        AppointmentID            = $apptID
        PatientFullNameLookup    = $patientLookupId  # References Patient's FullName calculated field
        DoctorFullNameLookup     = $doctorLookupId   # References Doctor's FullName calculated field
        AppointmentDateTime      = $date.ToString('yyyy-MM-ddTHH:mm:ss')
        AppointmentEndTime       = $date.AddMinutes(45).ToString('yyyy-MM-ddTHH:mm:ss')
        ServiceType              = $serviceType
        Status                   = if ($date -lt (Get-Date)) { @("Completed", "No Show", "Cancelled") | Get-Random } else { @("Scheduled", "Confirmed", "Rescheduled") | Get-Random }
        Notes                    = "Appointment notes for $($patient["Title"]) with $($doctor["Title"])."
        IsUrgent                 = $(if ((Get-Random -Minimum 1 -Maximum 10) -le 2) { $true } else { $false })
    }
}

$appointmentCount = 0
foreach ($appointment in $appointments) {
    try {
        Add-PnPListItem -List "Appointments" -Values $appointment | Out-Null
        $appointmentCount++
        Write-Host "  Added appointment: $($appointment.Title)" -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed to add appointment: $($appointment.Title) - $($_.Exception.Message)" -ForegroundColor Red
    }
}

# === GET APPOINTMENT IDs FOR PATIENT JOURNEY ACTIVITIES ===
Write-Host "Retrieving Appointment IDs for Patient Journey Activities..." -ForegroundColor Yellow

$appointments = Get-PnPListItem -List "Appointments"
Write-Host "  Found $($appointments.Count) appointments for activities" -ForegroundColor Cyan

# Create lookup maps for PatientJourneyActivities
$appointmentLookup = @{}
foreach ($appointment in $appointments) {
    if ($appointment["AppointmentID"]) {
        $appointmentLookup[$appointment["AppointmentID"]] = $appointment.Id
    }
}
Write-Host "  Created appointment lookup map: $($appointmentLookup.Count) appointments" -ForegroundColor Cyan

# === ACTIVITIES DATA GENERATION ===
$activityTypes = @("Admission", "Consultation", "Lab Test", "Medication", "Treatment", "Imaging", "Follow-up Call", "Discharge")
$priorities = @("Normal", "High", "Low", "Urgent")

$activities = @()
for ($i = 1; $i -le 50; $i++) {
    $appt = $appointments[($i-1) % $appointments.Count]
    $patient = $patients[($i-1) % $patients.Count]
    $doctor = $doctors[($i-1) % $doctors.Count]
   
    $activityID = Get-ActivityID $i
    $activityType = $activityTypes | Get-Random
    $priority = $priorities | Get-Random
   
    # Get lookup IDs for the relationships
    $appointmentLookupId = $appointmentLookup[$appt["AppointmentID"]]
    $patientLookupId = $patientLookup[$patient["PatientID"]]
    $doctorLookupId = $doctorLookup[$doctor["DoctorID"]]
   
    $activities += @{
        Title = "$activityType - $priority for $($patient["Title"])"
        ActivityID = $activityID
        AppointmentIDLookup = $appointmentLookupId
        PatientFullNameLookup = $patientLookupId    # Updated field name
        DoctorFullNameLookup = $doctorLookupId      # Updated field name
        ActivityDateTime = (Get-RandomDate -StartDate $startDate -EndDate $endDate).ToString("yyyy-MM-ddTHH:mm:ss")
        ActivityType = $activityType
        Notes = "$activityType notes for appointment $($appt["AppointmentID"])."
        Duration = @(15, 30, 45, 60) | Get-Random
        Priority = $priority
    }
}

$activityCount = 0
foreach ($activity in $activities) {
    try {
        Add-PnPListItem -List "PatientJourneyActivities" -Values $activity | Out-Null
        $activityCount++
        Write-Host "  Added activity: $($activity.Title)" -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed to add activity: $($activity.Title) - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "=== SAMPLE DATA IMPORT COMPLETE ===" -ForegroundColor Magenta
Write-Host "Successfully imported sample data:" -ForegroundColor Green
Write-Host "• $patientCount Patients with diverse medical conditions and demographics" -ForegroundColor Green
Write-Host "• $doctorCount Doctors across different medical specializations" -ForegroundColor Green
Write-Host "• $appointmentCount Appointments spanning May 15 - July 16, 2025" -ForegroundColor Green
Write-Host "• $activityCount Patient Journey Activities tracking care progression" -ForegroundColor Green
Write-Host ""
Write-Host "Date Range: May 15, 2025 to July 16, 2025" -ForegroundColor Cyan
Write-Host "The medical system is now populated with test data for validation!" -ForegroundColor Cyan
Write-Host "Data is optimized for Power Apps SharePoint lookup field compatibility." -ForegroundColor Cyan

# === VERIFICATION ===
Write-Host ""
Write-Host "=== DATA VERIFICATION ===" -ForegroundColor Magenta
$finalPatients = Get-PnPListItem -List "Patients"
$finalDoctors = Get-PnPListItem -List "Doctors"
$finalAppointments = Get-PnPListItem -List "Appointments"
$finalActivities = Get-PnPListItem -List "PatientJourneyActivities"

Write-Host "Final counts in SharePoint lists:" -ForegroundColor Cyan
Write-Host "• Patients: $($finalPatients.Count)" -ForegroundColor White
Write-Host "• Doctors: $($finalDoctors.Count)" -ForegroundColor White
Write-Host "• Appointments: $($finalAppointments.Count)" -ForegroundColor White
Write-Host "• Patient Journey Activities: $($finalActivities.Count)" -ForegroundColor White