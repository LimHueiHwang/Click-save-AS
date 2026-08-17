# Click Save As Automation

**Status: Production**

Python automation developed to improve the reliability of an existing SAP Purchase Order (PO) softcopy saving workflow. The solution uses OpenCV-based UI element detection instead of relying entirely on fixed screen coordinates.

## Overview

The automation replaces a fragile fixed-coordinate click with computer-vision-based detection of the SAP **Save As** UI element. The detected position is then used by PyAutoGUI to perform the required interaction as part of an existing Excel VBA and SAP GUI workflow.

## Problem

The existing desktop automation depended on fixed screen coordinates. Changes in screen resolution, display scaling, or application layout could cause the automation to click the wrong location or fail to locate the Save As function.

## Solution

The Python component captures the current screen and uses OpenCV template matching to locate the Save As UI element. It searches across multiple template scales and applies a configurable confidence threshold before clicking the detected position.

Both Dark Mode and Light Mode reference images are supported. Detection is retried when the target cannot be identified immediately.

## Workflow

![Workflow Diagram](docs/diagrams/workflow.png)

**Excel VBA / SAP workflow → Python → Screen capture → OpenCV template matching → PyAutoGUI click → Continue workflow**

## Architecture

![Architecture Diagram](docs/diagrams/architecture.png)

The Python component operates as part of the existing desktop automation rather than replacing the surrounding VBA and SAP workflow.

## Technologies

- Python
- OpenCV
- NumPy
- PyAutoGUI
- PyWin32
- openpyxl
- Excel VBA
- SAP GUI

## Key Features

- OpenCV template matching
- Multi-scale UI element detection
- Configurable matching threshold and scale range
- Dark Mode and Light Mode template support
- Retry handling for UI detection
- PO number normalization for Excel/VBA comparison
- Excel-based error reporting
- Windows user-profile-based local paths

## Error Handling

The automation retries UI detection up to the configured maximum number of attempts. When processing fails, the Python component records the error in an Excel error log and attempts to update the corresponding PO row in the running VBA workbook.

## Configuration

`utils/constant.example.py` provides the public configuration template.

Copy it to:

```text
utils/constant.py
```

and replace the example values with local settings.

The actual `constant.py` is excluded from Git through `.gitignore` so environment-specific paths and workbook names are not committed.

Runtime template images are expected in the user's local `PO Softcopy` folders. The images included in this repository are sanitized reference/sample assets for documentation and demonstration.

## My Role

I developed the Python component that detects the SAP Save As UI element and performs the required desktop interaction. I implemented the OpenCV matching logic, multi-scale detection, retry handling, configuration, PO number handling, and Excel error reporting used by the automation workflow.

## Business Impact

The solution addresses a production reliability problem in an SAP Purchasing automation workflow by reducing dependence on fixed screen coordinates and supporting different local display environments.

## Limitations

- Detection still depends on the expected appearance of the SAP/application UI.
- Runtime template images must be available in the configured local location.
- The solution is designed for the specific workflow and environment for which it was developed.
- No automated test suite is currently included.

## Future Improvements

- Package runtime template assets with the application.
- Add automated tests.
- Further separate application components as the automation grows.
- Improve UI detection robustness for additional interface variations.

## Disclaimer

This repository contains sanitized code and sample/reference assets for portfolio demonstration. Production files, company-specific data, credentials, and confidential information are not included.
