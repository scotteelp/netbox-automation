# Module Validation
import subprocess
import importlib

# Import color constants from color_definitions.py
from color_definitions import BOLD, UNDERLINE, RESET, RED, GREEN, YELLOW, BLACK, BLUE, MAGENTA, CYAN, WHITE, BG_BLACK, BG_RED, BG_GREEN, BG_YELLOW, BG_BLUE, BG_MAGENTA, BG_CYAN, BG_WHITE

# List of required modules
required_modules = ['pynetbox', 'csv', 'sys', 'requests', 'pandas', 'datetime', 'openpyxl']

# Check if a module is installed
def is_module_installed(module_name):
	try:
		importlib.import_module(module_name)
		return True
	except ImportError:
		return False

# Install a missing module
def install_module(module_name):
	try:
		subprocess.check_call(["pip3", "install", module_name])
		print(BOLD + BG_CYAN + WHITE +f"☑️  Successfully installed {module_name}" + RESET)
	except Exception as e:
		print(BOLD + UNDERLINE + RED + f"🅧  Error installing {module_name}: {e}" + RESET)

# Check and install missing modules
def check_and_install_modules(module_list):
	missing_modules = [module for module in module_list if not is_module_installed(module)]
	
	if missing_modules:
		print("⚠️  The following required modules are missing:")
		for module in missing_modules:
			print(f" - {module}")
		
		install_choice = input("Do you want to install the missing modules? (y/n): ").lower()
		
		if install_choice == "y":
			for module in missing_modules:
				install_module(module)
	
	else:
		print(BOLD + BG_CYAN + WHITE + "☑️  All required modules are already installed." + RESET)

# Call the function to check and install modules
check_and_install_modules(required_modules)

# Modules
import os
import pynetbox
import json
import csv
import sys
from csv import writer
import random
import requests
import pandas as pd
import datetime
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment

# Load sensitive data from config.py and store as environment variables
from config import NETBOX_TOKEN, NETBOX_URL
os.environ["NETBOX_TOKEN"] = NETBOX_TOKEN
os.environ["NETBOX_URL"] = NETBOX_URL

# Import ascii art from ascii_art.py
from ascii_art import VIDGO_ASCII, FACE_ASCII, CHUCK_ASCII, NETBOX_ASCII

def get_freewheel_data():
	url = "https://api.freewheel.tv/services/v4/sites/1"  # Freehweel API endpoint

	headers = {
		"Authorization": "Bearer ACCESS_TOKEN", 
		"Accept": "application/json"
	}

	response = requests.get(url, headers=headers)

	if response.status_code == 200:
		data = response.json()
		# Process and utilize the data as needed
		print(BOLD + "FreeWheel TV Data:")
		print(BG_WHITE + BLACK + data)
	else:
		print(BOLD + "Failed to fetch FreeWheel TV data.")
		
#host and token
nb = pynetbox.api(NETBOX_URL, NETBOX_TOKEN)
#fetch all devices
nb_devicelist = nb.dcim.devices.all()

headers = ['Name', 'Status', 'Site', 'Rack', 'Role', 'Manufacturer', 'Type', 'Owner', 'Birthday', 'Age (Months)', 'Service Contract', 'Warranty', 'Serial Number', 'Platform', 'Software', 'SW_Version', 'Primary IP']

def calculate_age_in_months(birthday):
	today = datetime.today()
	birth_date = datetime.strptime(birthday, '%Y-%m-%d')
	age_months = (today.year - birth_date.year) * 12 + today.month - birth_date.month
	return age_months
	
def csv_to_xlsx(headers, devices_data):
	wb = openpyxl.Workbook()
	ws = wb.active

	ws.append(headers)
	for device in devices_data:
		ws.append([device.get(header, '') for header in headers])

	for col in ws.columns:
		max_length = 0
		column = col[0].column_letter  # Get the column name
		for cell in col:
			try:
				if len(str(cell.value)) > max_length:
					max_length = len(cell.value)
			except:
				pass
		adjusted_width = (max_length + 2)
		ws.column_dimensions[column].width = adjusted_width

	for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
		for cell in row:
			cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

	wb.save('output.xlsx')
	
def get_devices(nb_devicelist, headers):
	devices_data = []  # List to hold device information
	print()
	print(UNDERLINE + BG_GREEN + BLACK + "................................................" + RESET)
	print()
	print(GREEN + NETBOX_ASCII + RESET)
	# Redirect stdout to a null device to hide output
	#original_stdout = sys.stdout
	#sys.stdout = open('output.csv', 'a')
	str1 = 'Active'
	for nb_device in nb_devicelist:
			result = {}
			status = str(nb_device.status)
			if status == str1:
				result['Name'] = str(nb_device)
				result['Status'] = status
				result['Site'] = str(nb_device.site)
				result['Rack'] = str(nb_device.rack)
				result['Role'] = nb_device.device_role.name
				result['Manufacturer'] = nb_device.device_type.manufacturer.name
				result['Type'] = str(nb_device.device_type)
				result['Owner'] = nb_device.custom_fields.get('owner')
				result['Birthday'] = nb_device.custom_fields.get('Birthday')
				age = nb_device.custom_fields.get('age')
				if age is None and result['Birthday']:
					result['Age (Months)'] = calculate_age_in_months(result['Birthday'])
				else:
					result['Age (Months)'] = age
				result['Service Contract'] = nb_device.custom_fields.get('service_contract')
				result['Warranty'] = nb_device.custom_fields.get('warranty')
				result['Serial Number'] = str(nb_device.serial)
				result['Platform'] = str(nb_device.platform)
				result['SW'] = nb_device.custom_fields.get('SW')
				result['SW_Version'] = nb_device.custom_fields.get('SW_Version')
				result['Primary IP'] = str(nb_device.primary_ip)

				devices_data.append(result)

			with open('output.csv', 'a', newline='') as f_object:
				writer_object = writer(f_object)
				for device in devices_data:
					writer_object.writerow([device.get(header, '') for header in headers])

	csv_to_xlsx(headers, devices_data)
	
	#sys.stdout.close()
	#sys.stdout = original_stdout  # Restore original stdout
	print(BG_GREEN + BLACK + "Netbox device information written to output.csv" + RESET)	
	print(UNDERLINE + "................................................" + RESET)
	print()

def update_age(nb_devicelist):
	for nb_device in nb_devicelist:
		status = str(nb_device.status)
		if status == "Active":
			birthday = nb_device.custom_fields.get('Birthday')
			if birthday:
				new_age = calculate_age_in_months(birthday)
				# Update the 'age' custom field in NetBox
				nb_device.custom_fields['age'] = new_age
				nb_device.save()
				print(f"Updated age for device {nb_device.name} to {new_age} months.")
				
def joke():
	response = requests.get("https://api.chucknorris.io/jokes/random")
	if response.status_code == 200:
		joke_data = response.json()
		joke_text = joke_data.get("value")
		print()
		print(YELLOW + CHUCK_ASCII + RESET)
		print()
		print(BG_YELLOW + BLACK + "Random Chuck Norris Joke:" + RESET)
		print(BG_YELLOW + BLACK + joke_text + RESET)
		print("................................................")
		print("________________________________________________")
	else:
		print("Failed to fetch a Chuck Norris joke.")
		print("................................................")
		print("________________________________________________")
		
def show_help():
	print()
	print(RED + VIDGO_ASCII + RESET)
	print(RED + FACE_ASCII + RESET)
	print(BOLD + "Available functions:" + RESET)
	print(" ► " + BG_GREEN + BLACK + "get_devices" + RESET + " or " + BG_GREEN + BLACK + "-d" + RESET + " ► GETS active device info from Netbox, writes output.csv and converts to output.xlsx file.")
	print(" ► " + BG_BLUE + WHITE + "update_age" + RESET + " or " + BG_BLUE + WHITE + "-a" + RESET + " ► This will update the age for all active devices on Netbox server.")
	print(" ► " + BG_WHITE + BLACK + "get_freewheel_data:" + RESET + " or " + BG_WHITE + BLACK + "-f" + RESET + " ► GETS data from Freewheel.")
	print(" ► " + BG_YELLOW + BLACK + "joke:" + RESET + " or " + BG_YELLOW + BLACK + "-j" + RESET +  " ► Prints random Chuck Norris joke.")
	print(BOLD + WHITE + " ► Usage: python netbox_api.py <function_name>")
	print(UNDERLINE + BG_CYAN + "................................................" + RESET)
	print()
	sys.exit(0)
	
def main():
	check_and_install_modules(required_modules)
	
	if len(sys.argv) < 2:
		show_help()

	function_name = sys.argv[1]

	if function_name == "get_devices" or function_name == "-d":
		get_devices(nb_devicelist, headers)
	elif function_name == "update_age"or function_name == "-a":
		update_age(nb_devicelist)
	elif function_name == "get_freewheel_data" or function_name == "-f":
			get_freewheel_data()
	elif function_name == "joke" or function_name == "-j":
		joke()
	elif function_name == "--help" or function_name == "-h":
		show_help()
	else:
		print(f"Function '{function_name}' not recognized.")

if __name__ == "__main__":
	main()