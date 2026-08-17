# Example configuration for Click Save As Automation
#
# Copy this file to:
#     utils/constant.py
#
# Then replace the example values with your local environment settings.

# Maximum number of attempts to locate the Save As button.
MAX_RETRIES = 10

# OpenCV template-matching confidence threshold.
MATCH_THRESHOLD = 0.8

# Multi-scale template matching range.
SCALE_MIN = 0.8
SCALE_MAX = 1.2
SCALE_STEPS = 10

# Enable OpenCV debug visualization.
DEBUG_MODE = False

# Name of the Excel workbook used by the automation.
target_wb_name = "YOUR_WORKBOOK_NAME.xlsm"

# Path to the Excel error log.
error_file = r"C:\Path\To\python_error_log.xlsx"
