import datetime
import os
import sys
import traceback

import cv2
import numpy as np
import pyautogui
import win32com.client as win32
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

from utils.constant import (
    MAX_RETRIES,
    error_file,
    target_wb_name,
)


# ============================================================
# CONFIGURATION
# ============================================================

MATCH_THRESHOLD = 0.8
SCALE_RANGE = np.linspace(0.8, 1.2, 10)
DEBUG_MODE = False


# ============================================================
# EXCEL ERROR WORKBOOK
# ============================================================

def close_error_file(file_path, retries=5):
    """
    Close the error workbook if it is currently open in Excel.

    The workbook is closed without saving because the error log
    is written separately by the automation.
    """
    if not file_path:
        return

    error_file_abs = os.path.abspath(file_path).lower()

    try:
        excel = win32.GetActiveObject("Excel.Application")
    except Exception:
        return

    for _ in range(retries):
        workbook_found = False

        try:
            workbooks = list(excel.Workbooks)
        except Exception:
            return

        for workbook in workbooks:
            try:
                workbook_path = os.path.abspath(workbook.FullName).lower()

                if workbook_path == error_file_abs:
                    workbook_found = True
                    workbook.Close(SaveChanges=False)
                    break

            except Exception:
                continue

        if not workbook_found:
            return


# ============================================================
# IMAGE RECOGNITION
# ============================================================

def find_and_click_icon_multiscale(
    image_path,
    threshold=MATCH_THRESHOLD,
):
    """
    Locate an icon on the current screen using multi-scale
    OpenCV template matching and click its center.

    Returns:
        True  - icon detected and clicked
        False - icon not found or detection failed
    """
    if not image_path or not os.path.exists(image_path):
        return False

    try:
        # Capture the current screen.
        screenshot = pyautogui.screenshot()

        screen_img = cv2.cvtColor(
            np.array(screenshot),
            cv2.COLOR_RGB2BGR,
        )

        gray_screen = cv2.cvtColor(
            screen_img,
            cv2.COLOR_BGR2GRAY,
        )

        # Load the reference icon.
        template = cv2.imread(image_path)

        if template is None:
            print(
                f"[Warning] Unable to load image: {image_path}"
            )
            return False

        template_gray = cv2.cvtColor(
            template,
            cv2.COLOR_BGR2GRAY,
        )

        template_height, template_width = template_gray.shape[:2]

        best_value = 0
        best_location = None
        best_scale = 1

        # Search the screen using multiple template scales.
        for scale in SCALE_RANGE:
            resized_template = cv2.resize(
                template_gray,
                None,
                fx=scale,
                fy=scale,
                interpolation=cv2.INTER_LINEAR,
            )

            resized_height, resized_width = resized_template.shape[:2]

            if (
                resized_height > gray_screen.shape[0]
                or resized_width > gray_screen.shape[1]
            ):
                continue

            result = cv2.matchTemplate(
                gray_screen,
                resized_template,
                cv2.TM_CCOEFF_NORMED,
            )

            _, max_value, _, max_location = cv2.minMaxLoc(result)

            if max_value > best_value:
                best_value = max_value
                best_location = max_location
                best_scale = scale

        # Check whether the best match meets the threshold.
        if best_value < threshold or best_location is None:
            print(
                f"🔍 Icon not found for "
                f"{os.path.basename(image_path)} | "
                f"Best confidence: {best_value:.2f}"
            )
            return False

        top_left = best_location

        bottom_right = (
            top_left[0] + int(template_width * best_scale),
            top_left[1] + int(template_height * best_scale),
        )

        center_x = (
            top_left[0]
            + (bottom_right[0] - top_left[0]) // 2
        )

        center_y = (
            top_left[1]
            + (bottom_right[1] - top_left[1]) // 2
        )

        # Optional debug preview.
        if DEBUG_MODE:
            debug_img = screen_img.copy()

            cv2.rectangle(
                debug_img,
                top_left,
                bottom_right,
                (0, 0, 255),
                2,
            )

            cv2.imshow("Detected Icon", debug_img)
            cv2.waitKey(0)
            cv2.destroyAllWindows()

        # Click the detected icon.
        pyautogui.moveTo(
            center_x,
            center_y,
            duration=0.3,
        )

        pyautogui.click()

        print(
            f"✅ Found icon {image_path} at "
            f"({center_x},{center_y}) | "
            f"Confidence: {best_value:.2f}"
        )

        return True

    except Exception as error:
        print(
            f"[Error] find_and_click_icon_multiscale: {error}"
        )
        return False


# ============================================================
# IMAGE LOCATION
# ============================================================

def find_mode_images():
    """
    Locate the Dark Mode and Light Mode Save As reference images.

    The automation checks both the local Desktop folder and
    the OneDrive Desktop folder.
    """
    user_profile = os.environ.get("USERPROFILE")

    if not user_profile:
        print(
            "[Error] USERPROFILE environment variable not found."
        )
        return None, None

    folders = [
        os.path.join(
            user_profile,
            "Desktop",
            "PO Softcopy",
        ),
        os.path.join(
            user_profile,
            "OneDrive - Jabil",
            "Desktop",
            "PO Softcopy",
        ),
    ]

    dark_path = None
    light_path = None

    for folder in folders:
        dark_candidate = os.path.join(
            folder,
            "Dark_Mode.png",
        )

        light_candidate = os.path.join(
            folder,
            "Light_Mode.png",
        )

        if os.path.exists(dark_candidate):
            dark_path = dark_candidate

        if os.path.exists(light_candidate):
            light_path = light_candidate

        if dark_path and light_path:
            break

    if not dark_path or not light_path:
        missing_files = []

        if not dark_path:
            missing_files.append("Dark_Mode.png")

        if not light_path:
            missing_files.append("Light_Mode.png")

        print(
            "[Error] Missing image file(s): "
            + ", ".join(missing_files)
        )

        return None, None

    return dark_path, light_path


# ============================================================
# SAP SAVE AS PROCESS
# ============================================================

def process_po(po_number):
    """
    Detect and click the SAP Save As icon.

    The automation checks Dark Mode and Light Mode reference
    images on each retry.

    The PO number is used for error reporting. The surrounding
    VBA workflow remains responsible for the overall PO process.
    """
    dark_image, light_image = find_mode_images()

    if not dark_image or not light_image:
        raise FileNotFoundError(
            "Unable to locate Dark Mode and Light Mode "
            "Save As reference images."
        )

    for attempt in range(1, MAX_RETRIES + 1):
        print(
            f"Attempt {attempt}/{MAX_RETRIES} "
            f"to locate SAP 'Save As' button..."
        )

        # Try Dark Mode reference image.
        if find_and_click_icon_multiscale(dark_image):
            return

        # Try Light Mode reference image.
        if find_and_click_icon_multiscale(light_image):
            return

    raise RuntimeError(
        "Failed to find SAP 'Save As' button "
        f"after {MAX_RETRIES} retries for PO {po_number}."
    )


# ============================================================
# ERROR LOGGING
# ============================================================

def log_error(po_number, error_text):
    """
    Write an automation error to the Excel error log.
    """
    if not error_file:
        return

    if not os.path.exists(error_file):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Errors"

        worksheet.append(
            ["PO Number", "Time", "Error"]
        )

        workbook.save(error_file)

    workbook = load_workbook(error_file)

    if "Errors" not in workbook.sheetnames:
        worksheet = workbook.create_sheet("Errors")

        worksheet.append(
            ["PO Number", "Time", "Error"]
        )
    else:
        worksheet = workbook["Errors"]

    time_str = datetime.datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    worksheet.append(
        [
            po_number,
            time_str,
            error_text,
        ]
    )

    last_row = worksheet.max_row

    worksheet[f"C{last_row}"].alignment = Alignment(
        wrap_text=True
    )

    workbook.save(error_file)


# ============================================================
# PO NUMBER
# ============================================================

def get_po_number():
    """
    Retrieve the PO number passed from the VBA macro.
    """
    if len(sys.argv) <= 1:
        raise ValueError("Unable to find PO number.")

    return sys.argv[1]


# ============================================================
# VBA WORKBOOK ERROR UPDATE
# ============================================================

def log_error_macro(po_number, error_text):
    """
    Write the error message to the corresponding PO row
    in the running Excel macro workbook.
    """
    try:
        excel = win32.GetActiveObject(
            "Excel.Application"
        )

        target_workbook = next(
            (
                workbook
                for workbook in excel.Workbooks
                if workbook.Name == target_wb_name
            ),
            None,
        )

        if not target_workbook:
            return

        worksheet = target_workbook.Sheets("Macro")

        row = 4

        while worksheet.Cells(row, 1).Value is not None:
            cell_po = int(
                float(
                    worksheet.Cells(row, 1).Value
                )
            )

            current_po = int(float(po_number))

            if cell_po == current_po:
                worksheet.Cells(row, 4).Value = error_text
                break

            row += 1

    except Exception as error:
        print(
            f"[Warning] Unable to update VBA workbook: {error}"
        )


# ============================================================
# MAIN
# ============================================================

def main():
    """
    Main entry point for the Save As automation.
    """
    close_error_file(error_file)

    try:
        po_number = get_po_number()

        try:
            process_po(po_number)

        except Exception as error:
            clean_error = (
                f"{type(error).__name__}: {error}"
            )

            print(clean_error)

            log_error(
                po_number,
                clean_error,
            )

            log_error_macro(
                po_number,
                clean_error,
            )

    except Exception:
        clean_traceback = traceback.format_exc()

        log_error(
            "UNKNOWN",
            clean_traceback,
        )


if __name__ == "__main__":
    main()