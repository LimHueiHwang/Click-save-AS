# SAP PO Softcopy Save As Automation

## Overview

This project improves the reliability of an existing SAP Purchase Order (PO) softcopy automation workflow by replacing fixed screen-coordinate clicking with **Python and OpenCV image-based detection**.

Python is used as a specialized component within the existing **Excel VBA + SAP GUI** workflow rather than replacing it.

## Problem

The original automation relied on fixed screen coordinates to locate the **Save As** controls. This approach can become unreliable when the SAP interface changes due to:

* Different screen resolutions
* Window positioning
* SAP theme changes
* UI scaling
* Different desktop environments

A failed click can interrupt the entire PO softcopy process.

## Solution

The automation captures the current screen and uses **OpenCV template matching** to locate the required Save As interface element.

The detection process:

1. Captures the current screen.
2. Loads the appropriate reference image.
3. Searches across multiple image scales.
4. Calculates the best matching location.
5. Validates the match against a confidence threshold.
6. Uses PyAutoGUI to perform the click.
7. Retries the detection when necessary.
8. Logs errors when the operation cannot be completed.

This provides a more flexible alternative to fixed screen coordinates.

## Key Features

* OpenCV template matching
* Multi-scale image detection
* Confidence-based match validation
* Automatic retry handling
* PyAutoGUI desktop interaction
* Excel VBA integration
* PO number validation and normalization
* Excel-based error logging
* Supports both Python and packaged EXE execution

## Architecture

```text
Excel VBA
    ↓
SAP PO / ME22N
    ↓
PO Softcopy / Save As
    ↓
Python EXE
    ↓
OpenCV Image Detection
    ↓
PyAutoGUI Click
    ↓
VBA Workflow Continues
```

## Technologies

* Python
* OpenCV
* NumPy
* PyAutoGUI
* PyWin32
* OpenPyXL
* Excel VBA
* SAP GUI Scripting

## Project Structure

```text
Click-save-AS/
├── Click_save_as.py
├── README.md
├── requirements.txt
├── .gitignore
├── utils/
│   └── constant.example.py
├── images/
│   ├── Dark_Mode.png
│   ├── Light_Mode.png
│   └── PO_softcopy_preview.png
└── docs/
    └── diagrams/
        ├── architecture.png
        └── workflow.png
```

## Error Handling

The automation includes retry logic and error logging to reduce workflow interruptions. When image detection fails, the process records the error and updates the existing Excel-based error reporting workflow.

## Business Impact

The solution improves the reliability of an existing SAP purchasing automation by reducing dependency on fixed screen coordinates while preserving the established VBA workflow.

It demonstrates practical integration of **SAP automation, Python, OpenCV, desktop automation, and Excel VBA** to solve a real business process problem.

## Limitations

The automation still depends on:

* SAP being in the expected workflow state
* Compatible SAP UI appearance
* Reference images being available
* Active desktop interaction
* Expected Excel workbook structure

This project is a targeted automation solution, not a general-purpose SAP automation framework.

## Project Status

**Status: Production Automation / Portfolio Project**

The Python refactoring and repository cleanup have been completed/in progress. Final EXE packaging and end-to-end validation remain part of the release process.
