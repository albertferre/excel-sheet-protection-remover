import zipfile
import os
import re


def remove_sheet_protection(xml_text):
    """
    Remove the sheetProtection element from the provided XML text using a regular expression.

    Parameters:
        xml_text (str): The XML content as a string.

    Returns:
        str: The modified XML content with the sheetProtection element removed.
    """
    # Define the regular expression pattern to match the <sheetProtection> element
    pattern = r"<sheetProtection[^>]*\s*/>"

    # Replace the matched pattern with an empty string to remove the element
    modified_xml_text = re.sub(pattern, "", str(xml_text), flags=re.IGNORECASE)

    return modified_xml_text


def process_zip_file(zip_file_path):
    """
    Process a zip file, remove sheet protection from worksheets, and save the modified contents to a new zip archive.

    Parameters:
        zip_file_path (str): The path to the zip file to process.
    """
    try:
        # Open the original zip file for reading
        with zipfile.ZipFile(zip_file_path, "r") as zfile:

            # Get the folder and file name
            folder_path = os.path.dirname(zip_file_path)
            new_file_name = os.path.basename(zip_file_path)
            new_file_name = "modified_" + new_file_name

            # Create a new zip file for writing the modified contents
            with zipfile.ZipFile(
                os.path.join(folder_path, new_file_name), "w"
            ) as modified_zfile:

                # Loop through each file in the original zip archive
                for file_info in zfile.infolist():
                    file_name = file_info.filename

                    # Read the content of the current file
                    with zfile.open(file_name) as file:
                        content = file.read().decode()

                    # Check if the file is in the 'xl\worksheets\' folder
                    worksheet_folder = "xl/worksheets/"
                    if file_name.startswith(worksheet_folder):
                        # Remove sheet protection from the content
                        modified_content = remove_sheet_protection(str(content))

                        # Save the modified content into the new zip archive
                        modified_zfile.writestr(file_name, modified_content)
                    else:
                        # If the file is not in the 'xl\worksheets\' folder, save it as it is
                        modified_zfile.writestr(file_name, content)

    except zipfile.BadZipFile as e:
        print(f"Failed to open zip file: {e}")
    except Exception as e:
        print(f"Error occurred: {e}")
