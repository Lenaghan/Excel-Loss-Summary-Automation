# Excel-Loss-Summary-Automation

### table_to_dictionary.bas
Extracts data from excel table with 16 columns and many rows. 
Creates a nested dictionary with a composite key as the primary key for the outer dictionary.
The composite key uses claim number and valuation date to allow for type 2 SCD.
The inner dictionary uses column names for keys.
While extracting, the text fields are normalized/standardized based on the field.
Functions include MakeSnakeCase, MakeProperCase, and MakeSentenceCase.
A unit test is included. 

1. ExtractTableToDictionary  
     Function returns a Scripting Dictionary representing data in a named table in excel.
     The data is extracted by looping through the table, standarizing values in the process before storing them.
     The inner loop includes the key, value pairs set by iterating through columns adding to th einner dictionary using the header as the key.
     The outer loop iterates through the rows of the table, adding each row to the outer dictionary using a composite key. 
2. MakeSnakeCase  
     Takes one string as input parameter, and returns one string as output.
     Uses regular expressions (regex) to transform the string to snake case.
3. MakeProperCase  
     Takes one string as input parameter, and returns one string as output.
     Uses an excel worksheet function to transform the string to proper case.
4. MakeSentenceCase  
     Takes one string as input parameter, and returns one string as output.
     Uses a loop and regex to transform the string into sentence case, removing uncommon special characters, as well. 
5. TestExtractTableToDictionary  
     Call the ExtractTableToDictionary function and sets a new dictionary with the return value.
     Uses nested loops to iterate through the dictionary.
     Prints the key, value pairs to the immediate window. 

Next: Aggregate claim counts and values by coverage type and year.   
Following: The end product will produce a presentable summary of the data in multiple pages, including aggregate values and details.
