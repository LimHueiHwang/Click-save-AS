# Example configuration for Click Save As Automation
#
# Copy this file to:
#     utils/constant.py
#
# Then replace the example values with your local environment settings.

# Maximum number of attempts to locate the Save As button.
MAX_RETRIES = 10

# Name of the Excel workbook used by the automation.
# Replace with the actual workbook name used in your environment.
target_wb_name = "YOUR_WORKBOOK_NAME.xlsm"

# Path to the Excel error log.
# Replace with a valid local or network path in your environment.
error_file = r"C:\Path\To\python_error_log.xlsx"
