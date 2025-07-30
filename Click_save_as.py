import pyautogui
import time
import os
import traceback
import datetime
import win32com.client
import win32com.client as win32
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
import sys

def close_error_file(error_file):

    if not error_file:
        return

    excel = win32.Dispatch("Excel.Application")
    for wb in excel.Workbooks:
        if os.path.abspath(wb.FullName).lower() == os.path.abspath(error_file).lower():
            try:
                wb.Close(SaveChanges=False)
                return
            except Exception as e:
                raise Exception(f"Failed to close work book: {e}")

def find_and_click_button(image_path, confidence_level=0.6):
    try:
        location = pyautogui.locateOnScreen(image_path, confidence=confidence_level)
        if location:
            center = pyautogui.center(location)
            pyautogui.moveTo(center.x, center.y, duration=0.4)
            pyautogui.click()
            return True
        else:
            return False
    except Exception:
        raise

def find_mode_images():
    try:
        user_profile = os.environ.get("USERPROFILE")
        if not user_profile:
            raise EnvironmentError("USERPROFILE environment variable not found.")

        folders = [
            os.path.join(user_profile, "Desktop", "PO Softcopy"),
            os.path.join(user_profile, "OneDrive - Jabil", "Desktop", "PO Softcopy")
        ]

        dark_path = None
        light_path = None

        for folder in folders:
            try:
                dark_candidate = os.path.join(folder, "Dark_Mode.png")
                light_candidate = os.path.join(folder, "Light_Mode.png")

                if os.path.exists(dark_candidate):
                    dark_path = dark_candidate
                if os.path.exists(light_candidate):
                    light_path = light_candidate

                if dark_path and light_path:
                    break
            except Exception as folder_error:
                print(f"[Warning] Failed to check folder '{folder}': {folder_error}")

        if not dark_path or not light_path:
            missing = []
            if not dark_path:
                missing.append("Dark_Mode.png")
            if not light_path:
                missing.append("Light_Mode.png")
            raise FileNotFoundError(f"Missing image file(s): {', '.join(missing)}")

        return dark_path, light_path

    except Exception as e:
        print(f"[Error] Unable to locate mode images: {e}")
        return None, None


def process_po(po_number):

    dark_img, light_img = find_mode_images()

    max_retries = 10

    for attempt in range(max_retries):
        try:
            if find_and_click_button(dark_img, 0.8):
                return
        except Exception:
            pass

        try:
            if find_and_click_button(light_img, 0.8):
                return
        except Exception:
            pass

        time.sleep(1)

    # If neither image was found after all retries
    raise Exception(f"Failed to find SAP 'Save As' button after {max_retries} retries for PO {po_number}")


def log_error(po_number, error_text):
    global error_file
    if not os.path.exists(error_file):
        wb = Workbook()
        ws = wb.active
        ws.title = "Errors"
        ws.append(["PO Number", "Time", "Error"])
        wb.save(error_file)
    wb = load_workbook(error_file)
    ws = wb["Errors"]
    time_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    ws.append([po_number, time_str, error_text])
    last_row = ws.max_row
    ws[f"C{last_row}"].alignment = Alignment(wrap_text=True)
    wb.save(error_file)

def get_po_number():
    if len(sys.argv)>1:
        po_number = sys.argv[1]
        return po_number
    else:
        raise Exception(f"Unable to find PO number")


def log_error_macro(po_number, error_text):

    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        target_wb = next((wb for wb in excel.Workbooks if wb.Name == target_wb_name), None)
        if not target_wb:
            return
        ws = target_wb.Sheets("Macro")

        row = 4
        while ws.Cells(row, 1).Value is not None:
            cell_po = int(float(ws.Cells(row, 1).Value))
            po_number = int(float(po_number))
            if cell_po == po_number:
                ws.Cells(row, 4).Value = error_text
                break
            row += 1
    except Exception:
        pass

if __name__ == "__main__":
    target_wb_name = "PO Softcopy creation Imac Version 1.7 - for Excel 2016(SAP 770) - 07242025.xlsm"
    error_file = r"\\sgsind0nsifsv01a\IMAC\MACROS\test\Po Softcopy\python error PO softcopy.xlsx"

    close_error_file(error_file)
    try:
        po_number = get_po_number()

        try:
            process_po(po_number)
        except Exception as e:
            clean_traceback = f"{type(e).__name__}: {e}"
            print(clean_traceback)
            # Log to separate error file
            log_error(po_number, clean_traceback)
            log_error_macro(po_number, clean_traceback)

    except Exception:
        clean_traceback = traceback.format_exc()
        log_error("UNKNOWN", clean_traceback)
