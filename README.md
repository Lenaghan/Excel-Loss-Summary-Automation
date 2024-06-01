# Excel-Loss-Summary-Automation

Extracts data from excel table with 16 columns and many rows. 
Creates a nested dictionary with a composite key as the primary key for the outer dictionary.
The composite key uses claim number and valuation date to allow for type 2 SCD.
The inner dictionary uses column names for keys.
While extracting, the text fields are normalized/standardized based on the field.
Functions include MakeSnakeCase, MakeProperCase, and MakeSentenceCase.
A unit test is included. 

Next: Aggregate claim counts and values by coverage type and year. 
Following: The end product will produce a presentable, well formatted summary of the data in multiple pages, including aggregate values and details on each claim.
