# TSP Project Launcher - Route Optimization and Scheduling Suite

A comprehensive desktop application suite for optimizing travel routes, clustering service regions, and scheduling appointments using the Traveling Salesman Problem (TSP) algorithms and geographic clustering.

You can download the app installer here: https://www.dropbox.com/scl/fi/7ke4svhpusabz9pmwk8kr/Smart_Scheduler_Setup_1_0.exe?rlkey=xqx9z03fa752bier1xlzn5sir&st=9nqp5yee&dl=1

## Overview

This project provides an integrated workflow for businesses that need to efficiently schedule and route field service appointments across multiple locations. The suite consists of four specialized applications accessed through a central launcher interface, all working with project-based data management.

## Features

### 1. Project Management

The launcher provides centralized project management capabilities:

- Create and manage multiple independent projects
- Each project stores its own locations, distances, clusters, and schedules
- Quick access to recent projects
- Organized file structure for all project data

### 2. Postcode Distance Calculator

Calculates travel distances between all location pairs using real-world driving data.

**Key Features:**
- Imports location data from CSV files
- Fetches actual driving distances and times via API
- Generates comprehensive distance matrices
- Creates both full matrices and pairwise distance lists
- Progress tracking for long calculations
- Automatic data caching for project reuse

**Outputs:**
- `distances.csv` - Pairwise distances between all locations
- `distance_matrix.csv` - Full matrix format for optimization algorithms

### 3. TSP Regional Clustering Optimizer

Groups locations into optimal service regions using hierarchical clustering and geographic analysis.

**Key Features:**
- Configurable number of regions
- Depot/home base selection for route optimization
- Service time and work hours constraints
- Visual map display of clustered regions with color coding
- Interactive region naming and color assignment
- Workload balancing across regions
- Convex hull visualization for geographic boundaries
- Export region assignments for scheduling

**Analysis Outputs:**
- Region summaries with customer counts and estimated times
- Visual clustering maps with color-coded regions
- Custom region naming with Outlook calendar color integration
- `clustered_regions.csv` - Location assignments to regions
- `region_summary.csv` - Statistics per region
- `region_names.csv` - Custom region labels and colors

### 4. Calendar Organizer

Assigns service regions to specific dates and manages the scheduling calendar.

**Key Features:**
- Interactive calendar interface
- Drag-and-drop region assignment to dates
- Visual region selection with customer counts
- Monthly calendar view with region color coding
- Schedule persistence and reloading
- Microsoft Outlook integration for calendar export
- Automatic appointment creation with region-based categories

**Functionality:**
- Load clustered regions from previous step
- Assign regions to work dates
- Save and reload scheduling plans
- Export schedule directly to Outlook calendar
- Color-coded calendar appointments by region

### 5. Smart Scheduler

Detailed appointment scheduling and route optimization within assigned regions.

**Key Features:**
- Weekly timetable view with time slot management
- Visual appointment booking interface
- Real-time route efficiency validation
- Travel time calculations between appointments
- Conflict detection and prevention
- Distance-based scheduling optimization
- Home base integration for start/end travel
- Microsoft Outlook appointment creation

**Scheduling Capabilities:**
- 30-minute time slot granularity
- Configurable work hours (default 8 AM - 7 PM)
- Adjustable appointment durations
- Maximum appointments per day limits
- Route efficiency thresholds to prevent inefficient routing
- Visual travel time blocking on schedule
- CSV export of confirmed appointments

**Route Optimization:**
- Validates appointment sequences for efficiency
- Calculates actual travel times between locations
- Warns about routes exceeding efficiency thresholds
- Prevents scheduling conflicts
- Optimizes travel from home base

## Project Workflow

### Typical Usage Pattern

1. **Create Project**: Use the launcher to create a new project with a unique name
2. **Calculate Distances**: Import locations and generate distance matrices
3. **Cluster Regions**: Define optimal service regions based on geography and constraints
4. **Organize Calendar**: Assign regions to specific work dates
5. **Schedule Appointments**: Book individual appointments within scheduled regions
6. **Export to Outlook**: Sync appointments to Microsoft Outlook calendar

### Data Flow

```
locations.csv (input)
    |
    v
Postcode Distance Calculator
    |
    v
distances.csv + distance_matrix.csv
    |
    v
TSP Regional Clustering Optimizer
    |
    v
clustered_regions.csv + region_summary.csv + region_names.csv
    |
    v
Calendar Organizer
    |
    v
region_schedule.csv
    |
    v
Smart Scheduler
    |
    v
confirmed_appointments.csv
    |
    v
Microsoft Outlook Calendar (optional export)
```

## Project File Structure

Each project creates the following files:

- `locations.csv` - Input file with customer postcodes/locations
- `distances.csv` - Pairwise distance data
- `distance_matrix.csv` - Full distance matrix for algorithms
- `clustered_regions.csv` - Location-to-region assignments
- `region_summary.csv` - Statistics for each region
- `region_names.csv` - Custom region names and Outlook colors
- `region_schedule.csv` - Date-to-region calendar assignments
- `confirmed_appointments.csv` - Detailed appointment bookings

## Technical Requirements

### Dependencies

- Python 3.x
- tkinter (GUI framework)
- pandas (data manipulation)
- numpy (numerical operations)
- matplotlib (visualization)
- scikit-learn (clustering algorithms)
- scipy (spatial algorithms)
- shapely (geometric calculations)
- requests (API calls for distances)
- pywin32 (Microsoft Outlook integration)

### System Requirements

- Windows operating system (for Outlook integration)
- Microsoft Outlook installed (for calendar export features)
- Internet connection (for distance calculations)

## Input Data Format

### locations.csv

The primary input file for your project. This file is required to get started.

**Required Format:**

```
postcode,client_name
PE27 5BH,Ingrid Isbister
CM23 2JT,Lewis Lewisham
CM3 4NQ,Mark Mason
CM6 1AE,Omar O'Malley
CO15 3DT,Adam Adams
EN11 8FN,Bob Baker
IP33 1UZ,Charlie Carter
SG2 9XU,George Graham
```

**Important Notes:**

- **Header row is required**: The first row MUST be `postcode,client_name`
- **Postcode column must come first**: Always put postcode before client_name
- **UK postcode format**: Postcodes should be valid UK postcodes (e.g., PE27 5BH, CM23 2JT)
- **Client names are optional but recommended**: Names enable the Global Display Toggle feature
- **No quotes needed**: Data should not be wrapped in quotes
- **One entry per row**: Each location gets its own row

**Global Display Toggle Feature:**

Once your project is created and data is processed through the applications, you can toggle between:
- Displaying **customer names** (readable, business-friendly)
- Displaying **postcodes** (location-specific, for route planning)

This toggle applies **globally across all applications** (Postcode Distance App, TSP Clustering, Calendar Organizer, and Smart Scheduler) and persists across sessions.

**Example locations.csv file:**

```csv
postcode,client_name
CO15 3DT,Adam Adams
EN11 8FN,Bob Baker
IP33 1UZ,Charlie Carter
NN16 0EF,Duke Dutton
MK16 0AG,Euain Ewanson
PE21 7QR,Frankie Fritz
SG2 9XU,George Graham
PE9 1PJ,Hank Harrison
PE27 5BH,Ingrid Isbister
HP23 5BN,Jack Johnson
```

**If you only have postcodes (no names):**

You can still use the application, but the Global Display Toggle will not provide any visual difference. Consider adding customer names for better usability.

```csv
postcode,client_name
SW1A 1AA,
W1A 0AX,
EC1A 1BB,
```

**Acceptable postcode formats:**

- Standard UK: `PE27 5BH`, `CM23 2JT`, `SG2 9XU`
- All uppercase or mixed case: Both `PE27 5BH` and `pe27 5bh` are accepted
- The app will normalize postcodes to uppercase automatically

## Configuration Options

### Clustering Parameters

- Number of regions to create
- Depot/home base postcode
- Average service time per appointment
- Daily work hours available
- Region naming and color assignment

### Scheduling Parameters

- Work day start and end times
- Appointment duration (minutes)
- Maximum appointments per day
- Route efficiency threshold (percentage above optimal)
- Time slot interval (30 minutes default)

## Outlook Integration

The Calendar Organizer and Smart Scheduler can export appointments directly to Microsoft Outlook:

- Creates calendar appointments with location details
- Assigns category colors matching region assignments
- Includes travel time in appointment descriptions
- Sets appointment reminders and priorities
- Handles recurring region schedules

## Use Cases

- Field service companies scheduling technician visits
- Sales teams planning customer visits by territory
- Healthcare providers organizing home visits
- Delivery route planning and optimization
- Any business requiring geographic service area management

## Limitations

- Distance calculations require active internet connection
- Large datasets may require significant processing time
- Outlook integration requires Windows and Microsoft Outlook
- Distance API may have rate limits or usage costs
- Clustering assumes geographic proximity is primary factor

## Support and Troubleshooting

### Common Issues

**Distance calculation fails**: Check internet connection and API availability

**Clustering produces unbalanced regions**: Adjust service time or work hours parameters

**Outlook export doesn't work**: Ensure Microsoft Outlook is installed and configured

**Project files missing**: Verify project directory exists and has write permissions

### Data Validation

All applications include input validation and error handling to prevent data corruption or invalid operations.

## Future Enhancements

Potential areas for expansion:

- Multi-day route optimization across regions
- Historical data analysis and reporting
- Mobile companion app for field workers
- Real-time traffic integration
- Customer priority weighting
- Skills-based technician assignment
- Automated rescheduling for cancellations
