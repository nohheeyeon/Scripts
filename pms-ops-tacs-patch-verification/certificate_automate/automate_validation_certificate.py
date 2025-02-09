import os

from document_updater import DocumentUpdater
from excel_data_processor import ExcelDataProcessor
from patch_directory_handler import PatchDirectoryHandler

if __name__ == "__main__":
    base_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "../data")

    excel_processor = ExcelDataProcessor()
    file_manager = PatchDirectoryHandler(base_path)
    document_updater = DocumentUpdater(base_path, excel_processor, file_manager)

    document_updater.process_document()
