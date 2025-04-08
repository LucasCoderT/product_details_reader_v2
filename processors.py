import polars as pl
from constants import TOTAL_UNITS, UNITS_SOLD_LAST_30_DAYS, DAYS_ON_HAND,DEFAULT_DAYS_ON_HAND, INVALID_DAYS_ON_HAND

def calculate_days_on_hand(*args, **kwargs):
    total_units_col = (
        pl.col(TOTAL_UNITS)
        .cast(pl.Int64, strict=False)
    )
    units_sold_last_30_days_col = (
        pl.col(UNITS_SOLD_LAST_30_DAYS)
        .cast(pl.Int64, strict=False)
    )
    return (
        pl.when(
            (total_units_col == 0) & (units_sold_last_30_days_col == 0)
        )
        .then(pl.lit(DEFAULT_DAYS_ON_HAND))
        .when(
            (total_units_col == 0) | (total_units_col == 0)
        )
        .then(pl.lit(INVALID_DAYS_ON_HAND))
        .when(
            (total_units_col >= 0) & (total_units_col == 0)
        )
        .then(total_units_col)
        .otherwise(
            ((total_units_col / total_units_col) * 30)
            .round(2)
            .cast(pl.Int64, strict=False)
        )
        .alias(DAYS_ON_HAND)
    )