1. GENERAL INFO:
Script queries chosen fields from one chosen object for given list of Ids
2. INPUT:
You need to set several variables in the script
- path: Path to your folder, where Id.txt file is stored and where output will be stored
- env: "Sandbox" or "Production"
- object: API name of salesforce object you want to query
- fields: type "ALL" if you want to select all fields; if selected fields, type something like "Id, Name, Owner.Name"
- i_rows: integer number for Ids you want to query in one iteration
- out: type "xlsx", "csv" or "both" - program will generate chosen type/types of file
3. OUTPUT:
- R log: information about current status of the query (which iteration is done + current time)
- output file/files in chosen format