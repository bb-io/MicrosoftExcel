# Blackbird.io Microsoft Excel

Blackbird is the new automation backbone for the language technology industry. Blackbird provides enterprise-scale automation and orchestration with a simple no-code/low-code platform. Blackbird enables ambitious organizations to identify, vet and automate as many processes as possible. Not just localization workflows, but any business and IT process. This repository represents an application that is deployable on Blackbird and usable inside the workflow editor.

## Introduction

<!-- begin docs -->

Microsoft Excel is a spreadsheet software that enables users to organize, analyze, and visualize data using tabular grids. It offers a range of features, including formulas, charts, and formatting options, making it a versatile tool for various data-related tasks and business applications.

## Exporting glossaries

To utilize the **Export glossary** action, ensure that the Excel worksheet mirrors the structure obtained from the **Import glossary** action result. Follow these guidelines:

- **Worksheet Structure**:
   - The first row serves as column names, representing properties of the glossary entity: _ID_, _Definition_, _Subject field_, _Notes_, _Term (language code)_, _Variations (language code)_, _Notes (language code)_.
   - Include columns for each language present in the glossary. For instance, if the glossary includes English and Spanish, the column names will be: _ID_, _Definition_, _Subject field_, _Notes_, _Term (en)_, _Variations (en)_, _Notes (en)_, _Term (es)_, _Variations (es)_, _Notes (es)_.
- **Optional Fields**:
    - _Definition_, _Subject field_, _Notes_, _Variations (language code)_, _Notes (language code)_ are optional and can be left empty.
- **Main term and synonyms**:
    - _Term (language code)_ represents the primary term in the specified language for the glossary.
    - _Variations (language code)_ includes synonymous values for the term.
- **Notes handing**:
    - Notes in the _Notes_ column should be separated by ';' if there are multiple notes for a given entry.
- **Variations Handling**:
    - Variations in the _Variations (language code)_ column should be separated by ';' if there are multiple variations for a given term.
- **Terms notes format**:
    - Each note in the _Notes (language code)_ column should follow this structure: **Term or variation: note**.
    - Notes for terms should be separated by ';;'. For example, 'money: may refer to physical or banked currency;; cash: refers to physical currency.'

## Feedback

Do you want to use this app or do you have feedback on our implementation? Reach out to us using the [established channels](https://www.blackbird.io/) or create an issue.

<!-- end docs -->
