# User input
# # select ADOM
adom = 'root'  # name of adom; if ADOM's are not in use, specify 'root'

# # specify filepath to MetaVariable Excel file
# file_path = "example.xlsx"
file_path = "path/to/your/file"

# # specify which sheet to use
# active_sheet = "metavariables"  # name of the sheet containing metavariables
active_sheet = None  # specify None to use the active (first) sheet of the Excel file

# # select devices - select the devices to include
# device_list = "Nashville_Spoke_01"  # include just this device
# device_list = ["Nashville_Spoke_01", "LasVegas_Spoke_02"]  # include any device with the hostname in list
device_list = None  # include all devices

# # toggle device negation
negate_devices = False  # set to true to include every device except the one(s) listed in the device_list

# # select MetaVariables - select the MetaVariables to include
# variable_list = "inet_a_intf"  # include just this MetaVariable
# variable_list = ["inet_a_intf", "inet_a_ip"]  # include just these MetaVariables
variable_list = None # include all MetaVariables

# # toggle MetaVariable negation
negate_variables = True  # set to true to include every MetaVariable except the one(s) listed in the device_list

# # update default values
update_defaults = True  # if set to false, the global values row will be ignored

# # specify filepath to MetaVariable JSON file that will be outputted to/inputted from
json_output_file = "output.json"  # created when main.py is run
json_input_file = "metadata_variables.json"  # imported when main-json2xlsx.py is run

# # add new variable/device to Excel sheet when running main-jsxon2xlsx.py
add_new_vars = False  # set to True to add new variables to your Excel spreadsheet when converting JSON to Excel
add_new_devices = False  # set to True to add new devices to your Excel spreadsheet when converting JSON to Excel
