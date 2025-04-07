from cenums import Files
from my_types import MappedCell
from processors import calculate_days_on_hand

OUTPUT_MAPPED_CELLS: list[MappedCell] = [
    {
        'column_name': 'Merchant SKU',
        'file_name': Files.RESTOCK_REPORT,
        'original_column_name': 'Merchant SKU',
    },
    {
        'column_name': 'ASIN',
        'file_name': Files.RESTOCK_REPORT,
        'original_column_name': 'ASIN',
    },
    {
        'column_name': 'Product Name',
        'file_name': Files.RESTOCK_REPORT,
        'original_column_name': 'Product Name',
    },
    {
        'column_name': 'Part Number',
        'file_name': Files.INVENTORY_FILE,
        'original_column_name': 'Part Number',
    },
    {
        'column_name': 'Primary Supplier',
        'file_name': Files.INVENTORY_FILE,
        'original_column_name': 'Primary Supplier',
    },
    {
        'column_name': 'Classification',
        'file_name': Files.INVENTORY_FILE,
        'original_column_name': 'Classification',
    },
    {
        'column_name': 'Units Sold Last 30 Days',
        'file_name': Files.RESTOCK_REPORT,
        'original_column_name': 'Units Sold Last 30 Days',
    },
    {
        'column_name': 'Total Units',
        'file_name': Files.RESTOCK_REPORT,
        'original_column_name': 'Total Units',
    },
    {
        'column_name': 'Days on Hand',
        'processor': calculate_days_on_hand,
    },
    {
        'column_name': 'Quantity Available',
        'file_name': Files.INVENTORY_FILE,
        'original_column_name': 'Quantity Available',
    },
    {
        'column_name': 'Cost',
        'file_name': Files.FEED_VIZOR_PROCESSOR,
        'original_column_name': 'Inventory Product Cost ($)',
    },
    {
        'column_name': 'MIN_PRICE',
        'file_name': Files.FEED_VIZOR_PROCESSOR,
        'original_column_name': 'Floor Price ($)',
    },
    {
        'column_name': 'CURRENT_PRICE',
        'file_name': Files.FEED_VIZOR_PROCESSOR,
        'original_column_name': 'Listing Current Price ($)',
        'post_processor': 1,
    },
    {
        'column_name': 'MAX_PRICE',
        'file_name': Files.FEED_VIZOR_PROCESSOR,
        'original_column_name': 'Ceiling Price ($)',
    },
]