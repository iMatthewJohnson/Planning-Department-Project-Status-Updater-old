
import os
import shutil
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


def update_comments(project_status_workbook, status_responses_workbook):
    ps_wb = load_workbook(filename=project_status_workbook)
    sr_wb = load_workbook(filename=status_responses_workbook)

    sr_ws = sr_wb.active

    # Mapping of status letters to the corresponding columns in the project status workbook
    status_to_col = {
        'A': 'L',
        'B': 'M',
        'C': 'N',
        'D': 'O',
        'E': 'P',
        'F': 'Q'
    }

    # Iterate over action ID and comments in the responses workbook
    for sr_row in sr_ws.iter_rows(min_row=2, values_only=True):
        action_id = sr_row[5].lower()
        comment = sr_row[6]
        status_letter = sr_row[7][0] if sr_row[7] else ''  # Get first char of status, check if not None
        found = False

        # Go through each sheet in the project status workbook
        if not found:
            for sheet in ps_wb.sheetnames:
                if sheet in ["Merge", "Full Plan", "OWNERS LIST"]:  # Skip these sheets
                    continue
                ps_ws = ps_wb[sheet]
                for ps_row in ps_ws.iter_rows(min_row=2):
                    # Look for the matching action ID
                    if ps_row[4].value and ps_row[4].value.lower() == action_id:
                        # Clear existing statuses
                        for col_letter in 'LMNOPQ':  # Columns representing statuses A-F
                            col_idx = column_index_from_string(col_letter) - 1  # Convert to 0-indexed
                            ps_row[col_idx].value = None

                        # Set new status
                        if status_letter in status_to_col:
                            new_status_col = status_to_col[status_letter]
                            # Convert to 0-indexed column number
                            col_idx = column_index_from_string(new_status_col) - 1
                            ps_row[col_idx].value = 'X'
                            found = True
                            break
                if found:
                    # Update comment in column J (index 9) if action ID is found
                    ps_row[10].value = comment
                    break

    ps_wb.save(filename=project_status_workbook)


def duplicate_workbook(file_path):
    file_dir, file_name = os.path.split(file_path)
    name, ext = os.path.splitext(file_name)
    old_file_path = os.path.join(file_dir, f"{name} (old){ext}")
    shutil.move(file_path, old_file_path)
    shutil.copy2(old_file_path, file_path) # Use copy2 to preserve metadata


# Main function to run the program
def run_sync(project_status_wb_path, status_responses_wb_path):
    duplicate_workbook(project_status_wb_path)
    update_comments(project_status_wb_path, status_responses_wb_path)