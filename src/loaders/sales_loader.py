# src/loaders/sales_loader.py
# load data sales connectivity dari beberapa source

import pandas as pd
from pathlib import Path
import sys


from config import (
    SOURCE_PRODA_SRCA, SRCA_PRODUCT_MAPPING,
    SOURCE_PRODB_SRCA, SRCA_REGION_MAPPING,
    SOURCE_PRODA_SRCB, SRCB_DIVISION_MAPPING,
    SOURCE_PRODB_SRCB, SOURCE_PRODC_SRCB,
    )


# loader data sales connectivity ---------------------------------------------

# ========================================================
# =========== LOADER SOURCE A PROD_A, PROD_C, PROD_B SEG_A
# ========================================================
# Data sales source A PROD_A & PROD_C SEG_A
def load_srca_proda(territory:str) -> dict[str, pd.DataFrame] :
    df = pd.read_csv(f"data/input/{SOURCE_PRODA_SRCA}", sep='\t',
                     header=0, encoding='utf-16')

    # data filter : territory, period dan produk
    mask_proda_sega = (
        (df['TERRITORY'] == territory) &
        (df['PRODUCT'].isin(SRCA_PRODUCT_MAPPING['PROD_A'])) #type:ignore
    )

    df_proda_sega = df[mask_proda_sega][['PERIOD', 'SALES_QTY']]
    df_proda_sega = df_proda_sega.rename(columns={'PERIOD': 'period', 'SALES_QTY': 'value'})

    mask_prodc_sega = (
        (df['TERRITORY'] == territory) &
        (df['PRODUCT'].isin(SRCA_PRODUCT_MAPPING['PROD_C'])) #type:ignore
    )
    df_prodc_sega = df[mask_prodc_sega][['PERIOD', 'SALES_QTY']]
    df_prodc_sega = df_prodc_sega.rename(columns={'PERIOD': 'period', 'SALES_QTY': 'value'})

    return {
        'proda_seg_a': df_proda_sega,
        'prodc_seg_a': df_prodc_sega
            }

# data source A Sales PROD_B SEG_A
def load_srca_prodb(territory:str) -> dict[str, pd.DataFrame]:
    df = pd.read_csv(f"data/input/{SOURCE_PRODB_SRCA}",sep='\t',
                     header=0, encoding='utf-16')

    mask_prodb_sega = (
        (df['TERRITORY_NEW'] == SRCA_REGION_MAPPING.get(territory,territory))
    )
    df_prodb_sega = df[mask_prodb_sega][['PERIOD','SALES_SPEED']]
    df_prodb_sega = df_prodb_sega.rename(columns={'PERIOD': 'period', 'SALES_SPEED': 'value'})

    return {'prodb_seg_a':df_prodb_sega}

# ========================================================
# ====== LOADER SOURCE B PROD_A, PROD_C, PROD_B SEG_B, SEG_C & SEG_D
# ========================================================

# PROD_A Source B
def load_srcb_proda (territory:str) -> dict[str, pd.DataFrame]:
    df = pd.read_csv(f"data/input/{SOURCE_PRODA_SRCB}",sep='\t',
                     header=0, encoding='utf-16')

    # iterate untuk setiap divisi dan conditional filtering
    result = {}
    for divisi, segment in SRCB_DIVISION_MAPPING.items():
        mask_srcb_proda = (
            (df['territory_new'] == territory)&
            (df['PRODUCT_NAME']=='PROD_A')&
            (df['division_new'] == divisi)
        )
        df_srcb_proda = df[mask_srcb_proda][['period','sales_qty']]
        df_srcb_proda = df_srcb_proda.rename(columns={'period': 'period', 'sales_qty': 'value'})

        result[f'proda_{segment.lower()}'] = df_srcb_proda

    return result


# PROD_B Source B
def load_srcb_prodb (territory:str) -> dict[str, pd.DataFrame]:
    df = pd.read_csv(f"data/input/{SOURCE_PRODB_SRCB}",sep='\t',
                     header=0, encoding='utf-16')

    # iterate untuk setiap divisi dan conditional filtering
    result = {}
    for divisi, segment in SRCB_DIVISION_MAPPING.items():
        mask_srcb_prodb = (
            (df['territory_new'] == territory)&
            (df['division_new'] == divisi)
        )
        df_srcb_prodb = df[mask_srcb_prodb][['period','sales_speed']]
        df_srcb_prodb = df_srcb_prodb.rename(columns={'period': 'period', 'sales_speed': 'value'})
        result[f'prodb_{segment.lower()}'] = df_srcb_prodb

    return result

# PROD_C Source B
def load_srcb_prodc (territory:str) -> dict[str, pd.DataFrame]:
    df = pd.read_csv(f"data/input/{SOURCE_PRODC_SRCB}",sep='\t',
                     header=0, encoding='utf-16')

    # iterate untuk setiap divisi dan conditional filtering
    result = {}
    for divisi, segment in SRCB_DIVISION_MAPPING.items():
        mask_srcb_prodc = (
            (df['territory_new'] == territory)&
            (df['PRODUCT_NAME']=='PROD_C')&
            (df['division_new'] == divisi)
        )
        df_srcb_prodc = df[mask_srcb_prodc][['period','sales_qty']]
        df_srcb_prodc = df_srcb_prodc.rename(columns={'period': 'period', 'sales_qty': 'value'})
        result[f'prodc_{segment.lower()}'] = df_srcb_prodc

    return result
