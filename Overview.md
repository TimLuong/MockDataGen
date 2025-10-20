# CreateAndInsertData.ps1 - Overview

## Purpose
This PowerShell script is a comprehensive medical system data generator designed to populate SharePoint lists with realistic sample data for testing and demonstration purposes. It creates a complete healthcare management system with patients, doctors, appointments, and patient journey activities.

## Target Platform
- **SharePoint Online** (Microsoft 365)
- **SharePoint Site**: ProjectManagement site
- **Integration**: Optimized for Power Apps SharePoint lookup field compatibility

## Script Capabilities

### 1. Automatic List Creation
The script can automatically create four SharePoint lists with proper schema if they don't exist:
- **Patients**
- **Doctors**
- **Appointments**
- **PatientJourneyActivities**

Each list is created with:
- Custom ID fields (unique, indexed)
- Calculated fields (e.g., FullName)
- Lookup relationships between lists
- Choice fields with predefined options
- Proper field configurations for Power Apps integration

### 2. Data Generation Features

#### Patients List (30 records)
- **Custom ID Format**: `MRN00001` - `MRN00030` (Medical Record Number)
- **Fields Generated**:
  - Personal information (FirstName, LastName, DateOfBirth, Gender)
  - Contact details (ContactNumber, EmailAddress, Address)
  - Medical information (MedicalHistorySummary)
  - Status tracking (PatientStatus)
  - Calculated FullName field
- **Demographics**: Randomized from pools of 20 first names and 20 last names
- **Age Range**: 18-50 years old

#### Doctors List (10 records)
- **Custom ID Format**: `DOC0001` - `DOC0010`
- **Fields Generated**:
  - Professional information (FirstName, LastName, Specialization, Department)
  - Contact details (ContactEmail)
  - Calculated FullName with "Dr." prefix
- **Specializations**: 10 medical specializations (Cardiology, Neurology, etc.)
- **Departments**: 10 hospital departments

#### Appointments List (30 records)
- **Custom ID Format**: `APP000001` - `APP000030`
- **Fields Generated**:
  - Scheduling (AppointmentDateTime, AppointmentEndTime)
  - Service type (10 options: Consultation, Follow-up, Lab Test, etc.)
  - Status tracking (Scheduled, Confirmed, Completed, etc.)
  - Urgency flag (IsUrgent)
  - Lookup relationships to Patients and Doctors
- **Date Range**: Next 30 days from script execution
- **Duration**: 45 minutes per appointment
- **Business Hours**: 8 AM - 5 PM in 15-minute intervals

#### Patient Journey Activities List (50 records)
- **Custom ID Format**: `ACT000001` - `ACT000050`
- **Fields Generated**:
  - Activity tracking (ActivityType, ActivityDateTime, Duration)
  - Priority levels (Normal, High, Low, Urgent)
  - Notes and descriptions
  - Lookup relationships to Patients, Doctors, and Appointments
- **Date Range**: May 15, 2025 - July 16, 2025
- **Activity Types**: 8 types (Admission, Consultation, Lab Test, etc.)

### 3. Data Relationships
The script implements a hierarchical data model:
```
Patients (base) ─┬─> Appointments
                 │
Doctors (base) ──┘
                 
Appointments ───> Patient Journey Activities
Patients ────────> Patient Journey Activities
Doctors ─────────> Patient Journey Activities
```

All relationships use **lookup fields** that reference calculated `FullName` fields for better readability in Power Apps.

## Script Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `SiteUrl` | String | ProjectManagement site URL | Target SharePoint site URL |
| `ClearExistingData` | Switch | False | Removes all existing data before import |
| `CreateListsIfMissing` | Switch | True | Auto-creates lists if they don't exist |

## Key Functions

### List Creation Functions
- `New-PatientsListIfMissing()` - Creates Patients list with schema
- `New-DoctorsListIfMissing()` - Creates Doctors list with schema
- `New-AppointmentsListIfMissing()` - Creates Appointments list with schema
- `New-PatientJourneyActivitiesListIfMissing()` - Creates PatientJourneyActivities list with schema

### Utility Functions
- `Test-LastCommand()` - Validates command success/failure
- `Get-RandomDate()` - Generates random dates within a range during business hours
- `Clear-ListData()` - Removes all items from a specified list
- `Get-PatientID()`, `Get-DoctorID()`, `Get-AppointmentID()`, `Get-ActivityID()` - Generate formatted IDs

### Core Process Functions
- `Initialize-SharePointListsIfNeeded()` - Orchestrates list creation
- Data generation loops for each entity type
- Lookup relationship mapping and ID retrieval

## Authentication
- Uses **PnP PowerShell** module
- **ClientId-based authentication**: `895298f8-3468-445f-8059-0d925de101fc`
- Connects via `Connect-PnPOnline`

## Workflow

### Phase 1: Setup
1. Connect to SharePoint site
2. Check and create lists if missing (removes existing lists for clean slate)
3. Clear existing data if `-ClearExistingData` flag is set

### Phase 2: Data Generation
1. **Patients** - Generate 30 patient records with demographics
2. **Doctors** - Generate 10 doctor records with specializations
3. **Retrieve IDs** - Create lookup maps for patients and doctors
4. **Appointments** - Generate 30 appointments with patient/doctor lookups
5. **Retrieve Appointment IDs** - Create lookup map for appointments
6. **Activities** - Generate 50 patient journey activities with all relationships

### Phase 3: Verification
- Count and display final record counts in each list
- Provide summary statistics

## Design Considerations

### Power Apps Compatibility
- Uses calculated fields for display names (e.g., FullName)
- Lookup fields reference user-friendly calculated fields
- Indexed fields for better performance
- Proper unique constraint enforcement on ID fields

### Data Realism
- Realistic medical conditions pool
- Business hour scheduling
- Appropriate service types and priorities
- Professional email formats
- Formatted phone numbers and addresses

### Error Handling
- Try-catch blocks around data insertion
- Validation of successful operations
- Colored console output for success/failure
- Graceful degradation if lists don't exist

### Performance Optimizations
- Batch processing with lookup maps
- Indexed fields on unique identifiers
- Indexed lookup fields for faster queries

## Output
The script provides color-coded console output:
- **Cyan**: Informational messages
- **Green**: Success messages
- **Yellow**: Warnings
- **Red**: Error messages
- **Magenta**: Section headers

## Use Cases
1. **Testing Power Apps** - Provides sample data for app development
2. **Demo Scenarios** - Realistic healthcare data for presentations
3. **Training** - Practice environment for users
4. **Development** - Test environment for SharePoint customizations
5. **Integration Testing** - Validate data flows and relationships

## Requirements
- **PnP PowerShell Module** (`PnP.PowerShell`)
- **SharePoint Online** access
- **Appropriate permissions** to create lists and add items
- **PowerShell 5.1+** or PowerShell Core 7+

## Maintenance Notes
- Script includes "FIXED VERSION" and "ENHANCED VERSION" notes indicating iterative improvements
- Designed to handle schema changes by removing and recreating lists
- Can be run multiple times (use `-ClearExistingData` to reset)

## Data Volume Summary
- **Total Records Created**: 120 items
  - 30 Patients
  - 10 Doctors
  - 30 Appointments
  - 50 Patient Journey Activities
- **Date Coverage**: ~2 months (May-July 2025)
- **Relationship Complexity**: 3 levels of lookup relationships

## Future Enhancement Opportunities
1. Parameterize record counts
2. Add more diverse medical data (diagnoses, medications, test results)
3. Include billing/insurance information
4. Add more realistic temporal relationships (follow-ups after initial visits)
5. Export capability to backup generated data
6. Validation of data integrity after insertion
7. Support for incremental updates without full recreation
