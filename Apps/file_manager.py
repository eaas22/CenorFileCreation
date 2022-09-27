import numpy as np
import pandas as pd


def file_creation(cockpit_input, logi_md, expedition_number):
    df_cockpit = pd.read_excel(cockpit_input, sheet_name=0, converters={'WMS Shipment': str, '12NC': int, 'DN Nr': int})
    df_logi = pd.read_excel(logi_md, sheet_name=0, header=1, converters={'12NC': str, 'PCE EAN': str})
    df_cockpit.index = np.arange(2, len(df_cockpit) + 2)
    df_logi.index = np.arange(2, len(df_logi) + 2)
    df_cockpit = df_cockpit[df_cockpit['WMS Shipment'] == expedition_number]
    query_12nc = df_cockpit['12NC'].astype(str).to_list()
    df_logi = df_logi[df_logi['12NC'].isin(query_12nc)].drop_duplicates()
    df_logi = df_logi.loc[:, ['Short Description', 'PCE EAN']]
    df_logi = df_logi.assign(short_descrip_2=df_logi['Short Description'].str.split(' ').str[0])
    df_cockpit = df_cockpit.loc[:, ['DN Nr', 'SAP PO Reference','12NC', 'DN QTY']]
    print('Cockpit matches: ')
    print(df_cockpit)
    print('Logistic Master Data matches: ')
    print(df_logi)
    cenor_sequence_data = {
        'Delivery': df_cockpit['DN Nr'].to_list(),
        'CUSTOMER_PO': df_cockpit['SAP PO Reference'].to_list(),
        # 'CUSOMFR_ID': [],
        'SKU_ID': df_cockpit['12NC'].astype('int64').to_list(),
        'QTY_ORDERED': df_cockpit['DN QTY'].to_list(),
        'DESCRIPTION': df_logi['Short Description'].to_list(),
        'FAN': df_logi['PCE EAN'].to_list(),
        'REFPROV': df_logi['short_descrip_2'].to_list(),
    }
    export_df = pd.DataFrame(data=cenor_sequence_data)
    export_df.insert(1, 'CUSOMFR_ID', "100137094")
    export_df.to_excel('order_cenor.xlsx', sheet_name='CENOR', index=False)
