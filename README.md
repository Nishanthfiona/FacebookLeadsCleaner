# Facebook Leads Cleaner Excel

## Overview

The **Facebook Leads Cleaner Excel** tool is designed to clean and enhance leads captured via Facebook forms. Facebook forms often allow unstructured data input, leading to inconsistent phone numbers and addresses. This tool streamlines the process of cleaning and organizing lead data, ensuring that phone numbers are formatted correctly and addresses are standardized. Additionally, it dynamically assigns salespersons to leads based on geographic location.

## Features

- **Phone Number Cleaning**: Removes country codes, symbols (such as `+`, `"`, etc.), and any non-numeric characters from phone numbers to standardize them. This helps in situations where Facebook forms lack validation, and users input inconsistent phone number formats.
  
- **Address Correction with Bing Maps API**:  
  - If a lead's address contains only a street, village, town, or district name instead of a state, the tool uses the **Bing Maps API** to look up the correct state (for Indian leads).
  - For leads outside of India, the tool retrieves the lead's country based on the address information.

- **Dynamic Salesperson Assignment**: Automatically assigns a salesperson to each lead based on the lead’s state. This ensures that leads are routed to the correct salesperson based on their geographic location.

- **One-Click Operation**: All processes are automated and executed with a single click of a button, making it user-friendly and efficient.

## How It Works

1. **Phone Number Cleaning**:  
   - The phone numbers in the leads sheet are cleaned using Excel formulas. These formulas remove any non-numeric characters, leaving only a valid phone number.
   
2. **Address Correction**:  
   - The tool parses the address provided by the lead. If the state information is missing for Indian leads, it calls the Bing Maps API to fetch the correct state. For leads outside of India, the tool returns the country based on the given address.
   
3. **Salesperson Assignment**:  
   - Once the state or country is identified, the tool dynamically assigns the correct salesperson for that region.

4. **VBA Implementation**:  
   - The core functionality, including API calls and salesperson assignment, is implemented in VBA for smooth and efficient performance. The phone number cleaning, however, is handled through Excel formulas.

## Usage

1. Open the Excel file and ensure that the lead data is correctly filled in the designated columns (phone number, address, etc.).
2. Click the **"Clean Leads"** button.
3. The tool will automatically:
   - Clean the phone numbers.
   - Correct the address information using the Bing Maps API.
   - Assign the correct salesperson based on the lead’s location.
4. Review the cleaned and organized data, which will be ready for further processing or sales follow-up.

## Prerequisites

- **Excel**: This tool is built in Excel and requires Excel to run.
- **Bing Maps API Key**: The tool uses the Bing Maps API to look up location information. You'll need to have a valid Bing Maps API key to use this feature. Ensure the API key is correctly integrated into the tool's VBA code.

## Installation

1. Download or clone the Excel file.
2. Open the file in Excel.
3. Enter your **Bing Maps API Key** in the VBA code where prompted.
4. Start using the tool by importing your leads data and clicking the button to clean and organize.

## Limitations

- The tool currently processes phone number cleaning using formulas, and any special or unique cases may need manual review.
- Requires an active internet connection to interact with the Bing Maps API for address correction.
