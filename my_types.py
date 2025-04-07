import typing
import polars as pl
from cenums import Files


class MappedCell(typing.TypedDict):
    """A dictionary that maps a column name to a cell value from another file"""
    column_name: str  # The name of the column
    file_name: typing.NotRequired[Files]  # The file it should be mapped from
    original_column_name: typing.NotRequired[str]  # The name of the column to get the value from
    processor: typing.NotRequired[typing.Callable[[typing.Dict[str, str]], None]]  # A function that processes the value of the cell
    validator: typing.NotRequired[typing.Callable[[typing.Any], bool]]  # A function that validates the value of the cell
    dtype: typing.NotRequired[pl.DataType]  # The type of the value of the cell