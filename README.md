# NetBox Automation Script

This script automates data retrieval from NetBox API and provides various functionalities for data processing.

## Overview

The NetBox Automation Script interacts with the NetBox API to retrieve device information, calculate device age, update age data in NetBox, fetch data from FreeWheel, and display Chuck Norris jokes. The script offers multiple features and functionalities to assist in managing and analyzing network device data.

## Features

- Connects to the NetBox API to retrieve and process device information.
- Calculates the age of devices based on birthdate information.
- Updates the age information for active devices in NetBox.
- Fetches data from FreeWheel API.
- Displays random Chuck Norris jokes for entertainment.

## Prerequisites

- Python 3.x installed.
- Required Python packages: `pynetbox`, `csv`, `sys`, `requests`, `pandas`, `datetime`, `openpyxl`.
- Access to a NetBox instance with API access.
- Access to the FreeWheel API.

## Getting Started

1. Clone this repository to your local machine.
2. Install the required dependencies using:
3. Create a `config.py` file and provide your NetBox API token and URL as `NETBOX_TOKEN` and `NETBOX_URL` respectively.
4. Run the script using: `python netbox_api.py <function_name>`.

## Available Functions

- `get_devices` (`-d`): Retrieves active device information from NetBox, writes to `output.csv`, and converts it to an `output.xlsx` file.
- `update_age` (`-a`): Updates the age of active devices in NetBox based on birthdate information.
- `get_freewheel_data` (`-f`): Fetches data from the FreeWheel API.
- `joke` (`-j`): Displays a random Chuck Norris joke.

## Usage

Run the script with a specified function name to perform the desired action. For example:
- To retrieve device information and generate CSV and Excel reports: `python netbox_api.py get_devices`
- To update age information for active devices: `python netbox_api.py update_age`
- To fetch data from FreeWheel API: `python netbox_api.py get_freewheel_data`
- To display a Chuck Norris joke: `python netbox_api.py joke`

## Contributing

Contributions are welcome! Please fork this repository and create a pull request with your enhancements.
