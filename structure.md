# Travel Budget Planner - Project Requirements

## inbox

all dates validated in format of year-month-day e.g. 2025 MAR 29
all times validated in 24 hour format with a mouseover noting that it is based off of local time

### instructions

add guidelines/general process/framework in filtering choices:

- Booking.com
- filter customer rating 7.5+
- less than 3km from city center or main activity
-

- corroborate reviews/rating by sorting by recent
- cross-reference with:
  - flyertalk
  - tripadvisor (only consider users with many reviews and older account age)

### trip settings

this sheet is unnecessary.

### lodging comparison

Include dynamic drop-down menus, filters, and pivot tables for ease of use.

at top, user input for number of travelers

user input of:

- retailer/booking site
- url
- name of hotel/hostel/inn/lodging
- date check in
- date check out
- number of beds and size of beds
- number of bathrooms
- number of kitchen
- square footage
- cost of parking (nearby street parking)
  (default values of 0 or null)

user inputs of total cost (make mouseover comment that this should integrate taxes, fees, and booking site discount):

- booking site discount? value is percentage (default 0%)
- spreadsheet calculates:
  - average cost per night,
  - cost per person per night,
  - total cost per person

user inputs post-booking incentives and cash back:

- Rakuten?
- Apple Pay?

allow users to rate each option of lodging with regard to:

- convenience to arrive and depart,
- distance to main activities, and an
- overall perceived value,
- x factor (mouseover comment with something like "x factor is relative to other options") from 1-5
  1 being the worst to 5 being the best
  take sum of these values and give each lodging a cumulative score
  calculate score per total cost

add a function that allows duplicating/inserting below of the lodging option with whatever values have already been input, leaving the rest blank

give user option to input overall ranking

### transportation

if flight, give user input option to add:

- airline
- departing and returning flight numbers
- date and time of departure
- date and time of return
- originating airport
- terminating airport
- total costs, including taxes and fees

if car rental,

- rental company
- date of start
- date of end
- originating facility
- terminating facility
- incidental costs:
  - snow chains
  - snow tires?
  - toll road transponder?
- total costs, including taxes and fees

allow users to rate each option of transportation with regard to:

- length of time/duration of travel
- convenience,
- total cost
- perceived value, and an
- x factor from 1-5 (relative to other options)
  (1 being the worst to 5 being the best)

add a function that allows duplicating/inserting below of the travel option with whatever values have already been input, leaving the rest blank

### trip settings:

calculate number of days using dates in B column, not E

## Overview

Create a Google Apps Script-based travel budget planner with comprehensive trip planning, budgeting, and visualization features.

## Core Features

- Multi-sheet organization for different planning aspects
- Protected formulas and data validation
- Custom menu system with tools and utilities
- Consistent styling and formatting across sheets
- Data import/export capabilities
- Budget optimization tools
- Currency conversion
- Timeline visualization
- Cost breakdown charts

## Technical Requirements

### Project Structure

Follow the established directory structure:

```
src/
├── config/
│   └── config.gs            # Global configuration and constants
├── core/
│   ├── menu.gs             # Menu creation and management
│   ├── protection.gs       # Sheet and range protection utilities
│   └── utilities.gs        # Common utility functions
├── sheets/
│   ├── basic/
│   │   ├── instructions.gs # Instructions sheet creation
│   │   ├── dataTables.gs  # Reference data tables
│   │   ├── tripSettings.gs # Trip settings sheet
│   │   └── lodging.gs     # Lodging comparison sheet
│   └── advanced/
│       ├── transportation.gs # Transportation options sheet
│       ├── activities.gs    # Activities planner sheet
│       └── dashboard.gs     # Summary dashboard sheet
└── main.gs                  # Main setup and initialization
```

### Dependencies

- Core modules must be loaded before sheet-specific modules
- Config module must be loaded first
- All sheet modules depend on utility functions

### Styling Guidelines

- Use consistent color scheme defined in config
- Apply standard formatting for headers, input cells, and formulas
- Implement protected ranges for formulas and reference data
- Use data validation for user inputs

### Implementation Requirements

1. Each module should be self-contained with clear dependencies
2. Use consistent error handling and validation
3. Implement proper sheet protection and data validation
4. Follow Google Apps Script best practices
5. Document all functions with JSDoc comments

## Sheet-Specific Requirements

[Details for each sheet module...]
