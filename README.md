# Excel-Loss-Summary-Automation

Extracts data from excel table with 16 columns and many rows. 
Creates a nested dictionary with claim numbers as the key for the outer dictionary.
The inner dictionary uses column names for keys.
While extracting, the text fields are normalized/standardized based on the field.
Functions include MakeSnakeCase, MakeProperCase, and MakeSentenceCase.
A unit test is included. 

Next: The outer dictionary will need a composite primary key to allow type 2 SCD.
Following: The end product will produce a presentable, well formatted summary of the data in multiple pages, including aggregate values and details on each claim.
