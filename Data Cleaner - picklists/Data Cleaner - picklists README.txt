1. GENERAL INFO:
Script reads data and checks if it's content is equal to available picklists values
2. INPUT:
- .xlsx file containing data to be checked for consistency against picklists values in salesforce; fields api names must be in the first row
You need to set several variables in the script
- path: Path to your folder
- input_file_name: Name of the input file (containing data to be checked)
- env: "Sandbox" or "Production"
- object: API name of salesforce object you are checking
3. OUTPUT:
- New folder created in given directory, containing two files:
  - "Picklists checked ... " file - file containing data from the input file, with information if values are consistent with picklist values in salesforse
  - "Picklists list" - file containing correct picklist values per each given field from Salesforce