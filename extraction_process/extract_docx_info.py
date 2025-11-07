import base64
import pathlib
import tika
from tika import (
    parser,
    unpack
) 
tika.initVM()


def extract_all_info(docx_file: str) -> dict:
    """Extract embedded files from given file path.
    
    Args:
        docx_file (str): Path to the docx file.

    Returns:
        dict: Extracted embedded files.
    """
    processed_parser = parser.from_file(docx_file)
    processed_unpack = unpack.from_file(docx_file)

    processed_data = {
        'file_name': docx_file,
        'metadata': processed_parser['metadata'],
        'content': processed_unpack['content'],
        'attachments': processed_unpack['attachments']
    }
        
    return processed_data


def format_embedded_files(processed_data: dict) -> dict:
    """Converts embedded files to string format if they are image or binary files.
    
    Args:
        processed_data (dic): Output results from extract_all_info function.
    Returns:

        dict: Converted sorted embedded files.
    """ 
    extensions = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff','.svg', '.emf', '.pdf', '.docx', '.xlsx', '.msg']
    formated_files = {}
    for key, value in processed_data["attachments"].items():
        file_extension = pathlib.Path(key).suffix.lower()
        if file_extension in extensions:
            encoded_content = base64.b64encode(value).decode('utf-8')
            formated_files[pathlib.Path(key).name] = encoded_content
        else:
            decoded_content = value.decode('utf-8', errors='ignore')
            formated_files[pathlib.Path(key).name] = decoded_content
    
    return formated_files


def combine_extracted_data(docx_file: str) -> dict:
    """Combine extracted metadata, content and sorted embedded files.
    
    Args:
        file_path (str): Path to the file.
    Returns:
        dict: Combined extracted data.
    """
    processesed_data = extract_all_info(docx_file)
    formated_files = format_embedded_files(processesed_data)

    combined_data = {
            'file_name': processesed_data['file_name'],
            'metadata': processesed_data['metadata'],
            'content': processesed_data['content'],
            'attachment': formated_files
        }

    return combined_data

