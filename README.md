# xlsx_to_json.MK2
This project is meant to assist in uploading mass amounts of metavariable data to FortiManager v7.2
or later.  The "main.py" scripts takes a specifically formatted Excel file and converts it to a JSON file
that FortiManager will recognize for importing Metadata.  Alternatively, the "main-json2xlsx.py" script
takes a JSON file, presumably downloaded from FortiManager, and updates/creates an Excel file.

Update the "user_settings.py" script with settings specific to your use case.  This is the only file
you need to edit.

Note: *FMG-xlsx_json-converter.zip* is a standalone version of this project that can be run without python.


## Setup
* Download python
    * The best way to download Python is from the official website: https://www.python.org/
    * Python can also be downloaded from terminal: *sudo apt-get install python3.8*
* Clone this repository from Github
    * Option1: clone via GitHub Desktop
    * Option2: clone via git bash using *git clone https://github.com/JoshuaHartling/xlsx_to_json.MK2.git*
* Open this Project in your favorite IDE.  (I recommend Pycharm!)
* Download the Python modules openpyxl and json either into your project's
virtual environment (recommended) or directly into python.
    * Either download modules using your IDE, or follow these instructions to download modules using
    pip: https://packaging.python.org/en/latest/guides/installing-using-pip-and-virtual-environments/
    
## What's Included
* *example.xlsx* is an example Excel file that shows how to format your metavariable data.
    * The first two columns should be your hostname and vdom assignment respectively.  Leave vdom entry
    blank to use global vdom.
        * You can put columns to the left of hostname/vdom columns for notes, but they will not be picked up
        by the script.
    * The first two rows should be metavariable name and its global value respectively.  Leave global entry
    blank to not use a global value.
* *example_metadata_variables.json* is an example file of the format FortiManager expects metavariable
data to be in.  It was pulled directly from FortiManager v7.2.3.
* *user_settings.py* contains user configuration parameters.
* *main.py* is the main script that should be run to convert your Excel file into the 
properly formatted JSON.
* *main-json2xlsx.py* is a script that imports a JSON file (usually downloaded from FortiManager) and
updates your Excel file.  It also creates a backup of the current Excel file before making 
changes to it.
* *example_output.json* is an example of the output of *main.py*.

## User Input
There are several parameters that you can set within the *user_settings.py* script,
though the only parameter you always must set is the file_path.
* **adom:** This is the ADOM on FortiManager that you wish to upload metadata to.  If ADOM's have
not been enabled on FortiManager, leave it as "root" as that is the default ADOM.  
* **file_path:** This variable is the directory path to your Excel file.  It is recommended
to set this as the absolute path.
* **active_sheet:** This is the sheet in your Excel File that holds your metadata.  Set this
if your metavariable data is not on the active (first) sheet of your Excel File.
* **device_list:** is a filter that can be set to only include specific devices.
The filter can be either a string or a list.  Leave it as 'None' to not use it.
* **negate:** is a boolean that when 'True' reverses the filter, turning it into a Blacklist
instead of a whitelist.  All devices not included in the filter will be added to the JSON file.
* **update_defaults:** when 'True' (default) the global values AKA default values of each variable
 will be included the JSON file.  Set to 'False' if you wish to ignore the default values.
* **json_output_file:** the name of the JSON file that will be created when *main.py*
is run.
* **json_input_file:** *(only applicable to main-json2xlsx.py)* the name of the JSON file
that will be imported when *main-json2xlsx.py* is run.
* **add_new_vars:** *(only applicable to main-json2xlsx.py)* when this boolean is true,
variables in the JSON file that are absent from the Excel file will be added to it.
* **add_new_devices:** *(only applicable to main-json2xlsx.py)* when this boolean is true,
devices in the JSON file that are absent from teh Excel file will be added to it.

## Upload to FortiManager
To upload the JSON file to FortiManger, log into FortiManager and navigate to the appropriate ADOM
(if ADOM's have been enabled).  Navigate to Policy & Objects > Object Configurations > Advanced >
Metadata Variables.  *Note: you may have to enable Metadata Variables using Feature Visibility
under Tools.*

Click on the 'More' dropdown menu, then select 'Import Metadata Variables'.  Either drag and drop
your *output.json* file or click 'Browse' to navigate to it.  Click 'Import'.

**Remember to always review your *output.json* file before uploading it to FortiManager.**
