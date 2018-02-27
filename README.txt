The following files take information contained in a spreadsheet of billed hours, and generates output reports based on whether individuals have been billed to a high enough level. The billing data fields
are mapped to fields in a fillable Adobe PF, and information from billing records as well as preset information about the client is allowed to flow into easily generated high scale reports.

- ExtractPDFFieldNames.bas - I added to a file found on the internet. It is designed to take a X by Y spreadsheet of information, and create a map in another worksheet of this information.
                              These fields are then mapped to the fields of a specific Adobe PDF template.
- FillForms - I added to a file found on the internet. on the basis of the map as established in ExtractPDFFieldNames.bas, the information contained in the workbook, the billing-based information, is allowed to flow into a fillable PDF form
- runFile.cls - controls how many of the records in the original worksheet are to be analyzed for this reporting program
- CheckBoxControl.bas - adds checkbox controls to specific rows and columns in the worksheet which contains the raw information.
- README.txt - this file