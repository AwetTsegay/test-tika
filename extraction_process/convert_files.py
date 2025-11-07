import json
import io
import pickle
import polars as pl

from extraction_process.extract_docx_info import (
    combine_extracted_data
)


def convert_to_json(combined_data: dict) -> str:
    """Creates Json file.
    
    Convert processed document outputs (Python objects) into json format files.

    Args:
        combined_data (dict): Combined extracted data from combine_extracted_data function.
    Returns:
        str: Converted json file path.
    """
    json_data = json.dumps(combined_data, indent=4)


    return json_data


def create_polars_df(combined_data: dict) -> pl.DataFrame:
    """Creates Polars DataFrame.

    Convert processed document outputs (Python objects) into polars DataFrame.

    Args:
        combined_data (dict): Combined extracted data from combine_extracted_data function.     
    Returns:
        pl.DataFrame: Converted polars DataFrame.
    """
    polars_data = pl.from_dicts(combined_data, strict=False)

    return polars_data


def create_parquet_file(combined_df: pl.DataFrame) -> str:
    """Creates Parquet file.

    Converts a polars DataFrame to a Parquet file.

    Args:
        combined_df (pl.DataFrame): Polars DataFrame from create_polars_df function.
    Returns:
        str: Converted parquet file path.
    """
    buffer = io.BytesIO()
    combined_df.write_parquet(buffer)
    parquet_data = buffer.getvalue()

    return parquet_data


def create_pickle_file(combined_data: dict) -> str:
    """Creates a Pickle file.
    
    Convert processed document outputs (Python objects) into pickle format files.

    Args:
        combined_data (dict): Combined extracted data from combine_extracted_data function.
    Returns:
        str: Converted pickle file path.
    """
    pickle_data = pickle.dumps(combined_data)

    return pickle_data


def process_document(document_path: str) -> dict:
    """Processes a document and converts it to multiple formats.

    Args:
        document_path (str): Path to the document.
    Returns:
        dict: Paths to the converted files.
    """
    combined_data = combine_extracted_data(document_path)
    json_data = convert_to_json(combined_data)
    polars_data = create_polars_df(combined_data)
    parquet_data = create_parquet_file(polars_data)
    pickle_data = create_pickle_file(combined_data)

    return {
        "json": json_data,
        "parquet": parquet_data,
        "pickle": pickle_data
    }

