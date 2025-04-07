import datetime
import logging
import pathlib
from typing import Callable, TypeVar

import polars
import polars as pl
import xlsxwriter
from xlsxwriter.worksheet import Worksheet

# Import constants
from constants import (
    LOGGING_FASEXCEL_TYPES_DTYPE,
    HIDDEN_SKU_COLUMN,
    SKU_SUBSTR,
    SHEET_NAME,
    CURRENT_PRICE,
    MIN_PRICE,
    MAX_PRICE,
    CP_COLOR_COL,
    COLOR_RED,
    COLOR_GREEN,
    COLOR_ORANGE,
    OUTPUT_XLSX_FILE, LOOKUP_COLUMN
)

from cenums import Files
from exceptions import NoSkuColumnsFound
from my_types import MappedCell
from output import OUTPUT_MAPPED_CELLS
from tkinter import ttk
FILES_DIR = pathlib.Path(__file__).parent / "files"

# disable polars dtype warning
logging.getLogger(LOGGING_FASEXCEL_TYPES_DTYPE).disabled = True

T = TypeVar("T")


def call_with_progress(func: Callable[..., T], progress_bar: ttk.Progressbar = None, *args, **kwargs) -> T:
    result = func(*args, **kwargs)
    if progress_bar:
        progress_bar.step(1)
    return result


def letter_to_number(letter: str) -> int:
    """
    Converts a letter to its corresponding column number (1-indexed).
    :param letter: The letter to convert
    :return: The corresponding column number
    """
    return ord(letter.upper()) - ord('A')

def load_file(fp: pathlib.Path) -> pl.DataFrame:
    """
    Loads a file from the files directory and returns it as a polars DataFrame.
    :param fp: The file path
    :return: A polars DataFrame
    """
    print(f"Reading {fp.name}")
    if not fp.exists():
        raise FileNotFoundError(f"File {fp} does not exist.")
    return pl.read_excel(fp)


def filter_dataframe(df: pl.DataFrame, values: list, *, name: str = None) -> pl.DataFrame:
    # Identify columns that contain the substring 'sku'
    columns = [col for col in df.columns if SKU_SUBSTR in col.lower()]
    if not columns:
        raise NoSkuColumnsFound(name, columns)

    # Create a new column \`__SKU__\` that captures the matched SKU
    df = df.with_columns(
        pl.coalesce(
            *[pl.when(pl.col(c).is_in(values)).then(pl.col(c)) for c in columns]
        ).alias(HIDDEN_SKU_COLUMN)
    )

    # Filter rows to include only matched SKUs
    return df.filter(pl.col(HIDDEN_SKU_COLUMN).is_not_null())

def create_output_dataframe(dataframes: dict, mapped_cells: list[MappedCell]) -> pl.DataFrame:
    # Rename "Merchant SKU" or "Part Number" to "SKU" before joining
    restock = dataframes[Files.RESTOCK_REPORT]
    inventory = dataframes[Files.INVENTORY_FILE]
    feed_visor = dataframes[Files.FEED_VIZOR_PROCESSOR]

    # Join on "SKU"
    merged = (
        restock
        .join(inventory, on=HIDDEN_SKU_COLUMN, how="left", suffix="_inv")
        .join(feed_visor, on=HIDDEN_SKU_COLUMN, how="left", suffix="_fv")
    )

    # Build column expressions from mapped_cells
    col_expressions = []
    for cell in mapped_cells:
        target_name = cell["column_name"]
        processor = cell.get("processor")
        if processor:
            # If there's a processor, evaluate it to get the column expression
            col_expressions.append(processor(merged))
            continue


        if "file_name" in cell:
            # Get the proper suffix given the file name
            file_name = cell.get("file_name")
            suffix = ""
            if file_name == Files.INVENTORY_FILE:
                suffix = "_inv"
            elif file_name == Files.FEED_VIZOR_PROCESSOR:
                suffix = "_fv"
            column_name = cell.get("original_column_name")
            if f"{column_name}{suffix}" in merged.columns:
                column_name = f"{column_name}{suffix}"
            col_expressions.append(pl.col(column_name).alias(target_name))
        else:
            # If there's no file_name, set it to None or handle processing as needed
            col_expressions.append(pl.lit(None).alias(target_name))

    # Construct a final dataframe with only the desired columns
    return merged.select(col_expressions)


def apply_current_price_color(df: pl.DataFrame) -> pl.DataFrame:
    """
    Applies color formatting to the 'Current Price' column based on its value.
    :param df: The DataFrame to format
    :return: The formatted DataFrame
    """
    """
    MIN_PRICE < CURRENT_PRICE < MAX_PRICE = green
    MIN_PRICE >= CURRENT_PRICE = red
    CURRENT_PRICE <= MAX_PRICE = orange
    """
    return df.with_columns(
        pl.when(
            (pl.col(MIN_PRICE) < pl.col(CURRENT_PRICE)) & (pl.col(CURRENT_PRICE) < pl.col(MAX_PRICE))
        )
        .then(pl.lit(COLOR_GREEN))
        .when(pl.col(MIN_PRICE) >= pl.col(CURRENT_PRICE))
        .then(pl.lit(COLOR_RED))
        .otherwise(pl.lit(COLOR_ORANGE))
        .alias(CP_COLOR_COL)
    )

def apply_current_price_color_styles_xlsxwriter(df: polars.DataFrame,
                                                workbook: xlsxwriter.Workbook,
                                                color_list: list[str]
                                                ) -> None:
    # Get the worksheet
    worksheet: Worksheet = workbook.get_worksheet_by_name("Sheet1")

    # Find column indexes (1-based) in ws
    price_col = df.columns.index(CURRENT_PRICE)

    # Define formats for red, green, and orange
    red_format = workbook.add_format({'bg_color': '#FF0000'})
    green_format = workbook.add_format({'bg_color': '#00FF00'})
    orange_format = workbook.add_format({'bg_color': '#FFA500'})

    # Loop through the DataFrame rows
    for i, (row, color_value) in enumerate(zip(df.iter_rows(named=True), color_list), start=1):
        if color_value == COLOR_RED:
            fill_format = red_format
        elif color_value == COLOR_GREEN:
            fill_format = green_format
        else:
            fill_format = orange_format

        worksheet.write(i, price_col, row[CURRENT_PRICE], fill_format)


def main(
        restock_report_path: pathlib.Path = FILES_DIR / f"{Files.RESTOCK_REPORT}.xlsx",
        inventory_file_path: pathlib.Path = FILES_DIR / f"{Files.INVENTORY_FILE}.xlsx",
        feed_visor_processor_path: pathlib.Path = FILES_DIR / f"{Files.FEED_VIZOR_PROCESSOR}.xlsx",
        *,
        progress_bar: ttk.Progressbar = None,
):
    restock_report = call_with_progress(load_file, progress_bar, restock_report_path)
    inventory_file = call_with_progress(load_file, progress_bar, inventory_file_path)
    feed_visor_processor = call_with_progress(load_file, progress_bar, feed_visor_processor_path)

    skus = restock_report[LOOKUP_COLUMN].to_list()


    # Filter inventory_file for rows that have any matching sku columns from the skus variable
    filtered_restock_report = call_with_progress(filter_dataframe, progress_bar, restock_report, skus, name=Files.RESTOCK_REPORT)
    filtered_inventory = call_with_progress(filter_dataframe, progress_bar, inventory_file, skus, name=Files.INVENTORY_FILE)
    filtered_feed_visor = call_with_progress(filter_dataframe, progress_bar, feed_visor_processor, skus, name=Files.FEED_VIZOR_PROCESSOR)


    # Create a dict of Files enum to DataFrame
    dataframes = {
        Files.RESTOCK_REPORT: filtered_restock_report,
        Files.INVENTORY_FILE: filtered_inventory,
        Files.FEED_VIZOR_PROCESSOR: filtered_feed_visor
    }

    output_dataframe = call_with_progress(create_output_dataframe, progress_bar, dataframes, mapped_cells=OUTPUT_MAPPED_CELLS)
    output_dataframe = call_with_progress(apply_current_price_color, progress_bar, output_dataframe)

    # Extract the __CP_COLOR__ column from the DataFrame
    # and convert it to a list of colors
    color_list = output_dataframe.select(CP_COLOR_COL).to_series().to_list()

    # Remove the __CP_COLOR__ column from the DataFrame
    output_dataframe = output_dataframe.drop(CP_COLOR_COL)

    # Save the output DataFrame to an Excel file
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
    output_file_path = f"output_{timestamp}.xlsx"
    print(f"Writing output to {output_file_path}")
    wb = xlsxwriter.Workbook(output_file_path)
    call_with_progress(output_dataframe.write_excel, progress_bar, wb)
    # Apply color formatting
    call_with_progress(apply_current_price_color_styles_xlsxwriter, progress_bar, output_dataframe, wb, color_list)
    wb.close()
    return output_file_path




if __name__ == "__main__":
    main()
