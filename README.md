# Word to PDF Batch Converter & Automated Page Processor ![PowerShell](https://img.shields.io/badge/PowerShell-%235391FE.svg?style=flat&logo=powershell&logoColor=white)

[English](README.md) | [繁體中文](README_zh.md) 

## Demo
<img width="1731" height="253" alt="image" src="https://github.com/user-attachments/assets/56d2c843-40b6-47da-8821-274b4f9717f0" />

## Introduction
This tool utilizes PowerShell to drive the Word COM interface, automating the conversion of `.docx` files to PDF. The workflow includes document cleanup (removing comments and hidden revisions), inserting a PBC stamp on the first page, and generating dynamic footers.

## Core Features
* **Document Cleanup & Standardization**:
    * Executes `Comments.DeleteAll()` to remove all comments.
    * Sets `ShowRevisionsAndComments = $false` to hide all tracked changes.
    * Automatically configures page margins (`BottomMargin: 15`, `FooterDistance: 16`).
* **First-Page Stamp Insertion**:
    * Inserts `pbc_stamp.png` (width set to 3.0cm) in the top-left corner of the first page.
    * Uses absolute positioning relative to the page (`Top: 10`, `Left: 20`).
* **Filename Parsing & Footer Generation**:
    * **Parsing Logic**: Extracts the string before the first underscore (`_`) in the filename to use as the title.
    * **Footer Layout**:
        * Left side: Static text `By Ariel Lin`.
        * Right side: Displays `[Parsed Title] P.[Auto Page Number]`.
        * Font: Times New Roman, Size 6.

## Logic Example
* **Input File**: `ProjectA_2026_v1.docx`
* **Process**:
    1.  Delete all comments and hide revision tracks.
    2.  Extract `ProjectA` as the unique identifier.
    3.  Place the PBC stamp in the top-left margin of the first page.
    4.  Insert footer: `By Ariel Lin` [Tab] `ProjectA P.1`.
* **Output File**: `ProjectA_2026_v1.pdf`

## Technical Implementation
* **Language**: PowerShell
* **Key Technical Points**:
    * **Word COM Object**: Utilizes `New-Object -ComObject Word.Application` for background processing.
    * **Unit Conversion**: Defines `$cmToPoints = 28.35` for precise coordinate mapping.
    * **Page Object Model**: Iterates through `Sections.Footers` and uses `Shapes.AddPicture` for layout manipulation.
    * **Field Codes**: Uses `wdFieldPage` (33) to insert dynamic page numbering.

## Execution Steps
1.  Place `pbc_stamp.png` in the same folder as the script.
2.  Place the Word documents (`.docx`) to be processed in the same folder.
3.  Run the PowerShell script.
4.  The processed PDF files will be generated in the same directory.
