# SAP PO Softcopy Save As Automation

![Production](https://img.shields.io/badge/Status-Production-success)
![Python](https://img.shields.io/badge/Python-Automation-blue)
![SAP](https://img.shields.io/badge/SAP-GUI%20Automation-blue)
![Excel VBA](https://img.shields.io/badge/Excel-VBA-green)
![OpenCV](https://img.shields.io/badge/OpenCV-Image%20Recognition-orange)

> Python automation component integrated with an existing Excel VBA and SAP Purchase Order workflow to improve the reliability of the Save As interaction.

---

## Overview

This project solves a reliability issue within an existing SAP PO Softcopy automation.

The original workflow used fixed screen coordinates to click the SAP **Save As** icon. Changes in monitor resolution, display scaling, SAP window position, or UI layout could cause the coordinate to point to the wrong location.

Python was introduced as a specialized component rather than replacing the existing VBA/SAP workflow.

The Python component uses **OpenCV template matching** to detect the actual Save As icon and **PyAutoGUI** to click its detected position.

---

## Problem

```text
Fixed X/Y Coordinate
        ↓
Click Save As
        ↓
Environment changes
        ↓
Potentially incorrect click
