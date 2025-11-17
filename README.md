# Blackbird.io Microsoft 365 Excel

Blackbird is the new automation backbone for the language technology industry. Blackbird provides enterprise-scale automation and orchestration with a simple no-code/low-code platform. Blackbird enables ambitious organizations to identify, vet and automate as many processes as possible. Not just localization workflows, but any business and IT process. This repository represents an application that is deployable on Blackbird and usable inside the workflow editor.

## Introduction

<!-- begin docs -->

Microsoft 365 Excel is a spreadsheet software that enables users to organize, analyze, and visualize data using tabular grids. It offers a range of features, including formulas, charts, and formatting options, making it a versatile tool for various data-related tasks and business applications.

## Actions

- **Add new sheet row**   Adds a new row to the first empty line of the sheet
- **Add new table row**   Add new table row
- **Create table**    Create table
- **Create worksheet**    Create worksheet
- **Download sheet CSV file**  Download CSV file
- **Export glossary**  Export glossary from Excel worksheet
- **Find sheet row**  Providing a column address and a value, return row number where said value is located
- **Get sheet cell**  Get cell by address
- **Get sheet range**  Get a specific range of rows and columns in a sheet
- **Get sheet row**   Get row by address
- **Get sheet used range**    Get used range in a sheet
- **Get table row**   Get table row
- **Import glossary**  Import glossary as Excel worksheet
- **List table rows**  List table rows
- **Update sheet cell**   Update cell by address
- **Update sheet row**    Update row by start address
- **Update table row**    Update table row
- **Download workbook PDF file** Download workbook PDF file

## Exporting glossaries

To utilize the **Export glossary** action, ensure that the Excel worksheet mirrors the structure obtained from the **Import glossary** action result. Follow these guidelines:

- **Worksheet structure**:
   - The first row serves as column names, representing properties of the glossary entity: _ID_, _Definition_, _Subject field_, _Notes_, _Term (language code)_, _Variations (language code)_, _Notes (language code)_.
   - Include columns for each language present in the glossary. For instance, if the glossary includes English and Spanish, the column names will be: _ID_, _Definition_, _Subject field_, _Notes_, _Term (en)_, _Variations (en)_, _Notes (en)_, _Term (es)_, _Variations (es)_, _Notes (es)_.
- **Optional fields**:
    - _Definition_, _Subject field_, _Notes_, _Variations (language code)_, _Notes (language code)_ are optional and can be left empty.
- **Main term and synonyms**:
    - _Term (language code)_ represents the primary term in the specified language for the glossary.
    - _Variations (language code)_ includes synonymous values for the term.
- **Notes handing**:
    - Notes in the _Notes_ column should be separated by ';' if there are multiple notes for a given entry.
- **Variations handling**:
    - Variations in the _Variations (language code)_ column should be separated by ';' if there are multiple variations for a given term.
- **Terms notes format**:
    - Each note in the _Notes (language code)_ column should follow this structure: **Term or variation: note**.
    - Notes for terms should be separated by ';;'. For example, 'money: may refer to physical or banked currency;; cash: refers to physical currency.'

## SharePoint documents support
In all actions you can specify optional input parameter "SharePoint site name" in order to use excel files from SharePoint site in "Workbook" parameter dropdown. 

![SharePointSupport1](/image/README/SharePointSupport1.png) ![SharePointSupport2](/image/README/SharePointSupport2.png)

## Feedback

Do you want to use this app or do you have feedback on our implementation? Reach out to us using the [established channels](https://www.blackbird.io/) or create an issue.

<!-- end docs -->
