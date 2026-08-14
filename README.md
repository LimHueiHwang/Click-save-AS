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

The detection process includes:

* Screen capture
* Reference image loading
* Multi-scale template matching
* Confidence-based validation
* Automated PyAutoGUI clicking
* Retry handling
* Error logging

## Key Features

* OpenCV template matching
* Multi-scale image detection
* Confidence-based match validation
* Automatic retry handling
* PyAutoGUI desktop interaction
* Excel VBA integration
* PO number validation and normalization
* Excel-based error logging
* Supports Python and packaged EXE execution

## Architecture

![Architecture](docs/diagrams/architecture.png)

## Workflow

![Workflow](docs/diagrams/workflow.png)

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

Python refactoring and repository cleanup are complete. Final EXE packaging and end-to-end validation are part of the release process.
