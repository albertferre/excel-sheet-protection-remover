import argparse
import logging
import os
from src.utils import process_zip_file


def setup_logging():
    # Set up logging configuration
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s]: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def main():
    # Configure logging
    setup_logging()

    # Parse command-line arguments
    parser = argparse.ArgumentParser(
        description="Remove sheet protection from a Excel file."
    )
    parser.add_argument("excel_file", help="Path to the Excel file to process")
    args = parser.parse_args()

    excel_file_path = args.excel_file

    # Check if the specified zip file exists
    if not os.path.exists(excel_file_path):
        logging.error("The specified Excel file does not exist.")
        return

    try:
        # Process the zip file to remove sheet protection
        process_zip_file(excel_file_path)
        logging.info(
            "Sheet protection removed successfully. Modified file saved as 'modified_{}'.".format(
                os.path.basename(excel_file_path)
            )
        )
    except Exception as e:
        # Log any error that occurs during the process
        logging.error(f"An error occurred: {e}")


if __name__ == "__main__":
    main()
