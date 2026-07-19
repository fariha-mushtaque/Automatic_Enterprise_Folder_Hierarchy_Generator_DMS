# Automatic_Enterprise_Folder_Hierarchy_Generator (for DMS)
An Excel-driven VBA automation tool that dynamically creates enterprise Document Management System (DMS) folder structures using a configurable hierarchy and template. The solution supports hierarchical inheritance, dynamic template folders, safe re-execution, and requires no code changes when the folder structure is updated.

## Overview

This project automates the creation of enterprise Document Management System (DMS) folder structures directly from an Excel workbook.

Instead of manually creating hundreds or thousands of folders, users only update the Excel sheets and run a single VBA macro.

The solution is metadata-driven, meaning folder structures are controlled entirely from Excel without modifying the VBA code.

## Features

✔ Dynamic hierarchy generation

✔ Supports up to 8 hierarchy levels

✔ Blank-cell hierarchy inheritance

✔ Dynamic template folders

✔ Creates template folders only under terminal (leaf) folders

✔ Existing folders are preserved

✔ Safe to execute multiple times

✔ Automatic Windows folder-name sanitization

✔ No hardcoded folder names

✔ Excel-driven configuration

## Folder Structure Generated

```text

Engineering Design Basis (Level 1)
│
├── Drawings (Level 2)
│      ├── Power & Utilities (Second last level
│      │      ├── Utilities-IV (Last level)
│      │      ├── STG
│      │      └── Grid
│      │
│      └── Fertilizers
│             ├── Urea-IV
│             ├── DAP & TFA
│             └── General
│
└── Specifications
       ├── Power & Utilities
       └── Fertilizers

```

## Excel Workbook Structure

The workbook contains four sheets.

1. Hierarchy Sheet

This sheet defines the main folder hierarchy.

Example


| Level1 | Level2 | Level3 | Level4 |
| :--- | :--- | :--- | :--- |
| Engineering |  |  |  |
| | Process |  |  |
| |  | Process Description  |  |
| |  |  | Main Reactions |
| Supply Chain |  |  |  |
| | Material Management |  |  |
| |  | P&U |  |
| |  | | Uty. |
| |  |  | STGs |
| |  | Fert |  |
| |  |  | A-Fert |
| |  |  | B-Fert |


## Rules

Maximum of 8 levels.
Blank cells inherit the previous parent.
Do not rearrange columns.
Add or remove folders by editing this sheet only.
No VBA changes are required.

This sheet defines folders that should be created inside every terminal (leaf) folder.

Example

| Parent Folder |	Child Folder |
| Power & Utilities |	Utilities-IV |
| Power & Utilities |	STG

## Can I add more folders?

Yes.

Simply add more rows.

Example:

| Parent Folder | Child Folder |
|---------------|--------------|
| Projects | Shutdown |
| Projects | Revamp |
| Projects | Commissioning |

The VBA automatically creates

```text
Projects
    Shutdown
    Revamp
    Commissioning
```
No code changes required.

## Can I add another Parent Folder?

Absolutely.

Example

| Parent Folder | Child Folder |
|---------------|--------------|
| QA Documents | Inspection Reports |
| QA Documents | Test Certificates |

Automatically creates

```text
QA Documents

Inspection Reports

Test Certificates

```

No VBA modification is needed.

Can Parent/Child have another level?

Current Version : No.

The Templates sheet currently supports:

```text
Parent Folder
    Child Folder
```

only.

If you need:

```text
Power & Utilities
    Utilities-IV
        Electrical
            Motors
```
that would require extending the VBA.


## How to Run

Open the template workbook (For your own workbook, import the .bas file).
Enable Macros.
Press
Alt + F8
Run
CreateFolderHierarchy
Select the destination folder.
Wait until the process completes.

Done!

## Workflow

Update Hierarchy Sheet

↓

Update Templates Sheet

↓

Run VBA Macro

↓

Select Destination Folder

↓

Folder Structure Generated


## Technical Details

Language

VBA

Platform

Microsoft Excel

Compatibility

Windows
Excel 2016+
Microsoft 365


## Current Limitations

- Supports up to 8 hierarchy levels.
- Templates support one Parent → Child level only.
- Existing folders are never deleted.
- Files are not copied.
- Folder names must be valid Windows folder names.
- Requires write permission on the destination folder.

## License

MIT License
