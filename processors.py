import polars as pl
from constants import TOTAL_UNITS, UNITS_SOLD_LAST_30_DAYS, DAYS_ON_HAND,DEFAULT_DAYS_ON_HAND, INVALID_DAYS_ON_HAND

def calculate_days_on_hand(*args, **kwargs):
    return (
        pl.when(
            (pl.col(TOTAL_UNITS) == 0) & (pl.col(UNITS_SOLD_LAST_30_DAYS) == 0)
        )
        .then(pl.lit(DEFAULT_DAYS_ON_HAND))
        .when(
            (pl.col(TOTAL_UNITS) == 0) | (pl.col(UNITS_SOLD_LAST_30_DAYS) == 0)
        )
        .then(pl.lit(INVALID_DAYS_ON_HAND))
        .when(
            (pl.col(TOTAL_UNITS) >= 0) & (pl.col(UNITS_SOLD_LAST_30_DAYS) == 0)
        )
        .then(pl.col(TOTAL_UNITS))
        .otherwise(
            ((pl.col(TOTAL_UNITS) / pl.col(UNITS_SOLD_LAST_30_DAYS)) * 30)
            .round(2)
            .cast(pl.Utf8)
        )
        .alias(DAYS_ON_HAND)
    )