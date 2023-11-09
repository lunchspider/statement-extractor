# This script is designed to convert bank statements from pdf to excel.
#
# It has been tweaked on HDFC Bank Credit Card statements,
# but in theory you can use it on any PDF document.
#
# The script depends on camelot-py,
# which can be installed using pip #
# pip install "camelot-py[cv]"


import os
import argparse
import camelot
import pandas as pd
import numpy as np
from collections import defaultdict


def handle_transaction(df: pd.DataFrame, start_index: int, end_index: int): 
    start_index += 2;
    info: str = df.loc[start_index][1]
    start_index += 1
    for i in range(start_index, end_index):
        row = df.loc[i];
        if(row[0].strip() == ''):
            info = row[1]
            continue
        df.at[i, 'deleted'] = False
        df.at[i,3] = info


def extract_df(path):
    # The default values from pdfminer are M = 2.0, W = 0.1 and L = 0.5
    laparams = {'char_margin': 2.0, 'word_margin': 0.2, 'line_margin': 1.0}

    # Extract all tables using the lattice algorithm
    lattice_tables = camelot.read_pdf(path,  
        pages='all', flavor='lattice', line_scale=50, layout_kwargs=laparams)

    # Extract bounding boxes
    regions = defaultdict(list)
    for table in lattice_tables:
        bbox = [table._bbox[i] for i in [0, 3, 2, 1]]
        regions[table.page].append(bbox)

    df = pd.DataFrame()

    # Extract tables using the stream algorithm
    for page, boxes in regions.items():
        areas = [','.join([str(int(x)) for x in box]) for box in boxes]
        stream_tables = camelot.read_pdf(path, pages=str(page),
            flavor='stream', table_areas=areas, row_tol=5, layout_kwargs=laparams)
        if(page == 5):
            for table in stream_tables:
                print(table.df)
                print('we are here')
        dataframes = [table.df for table in stream_tables]
        dataframes = pd.concat(dataframes)
        df = df._append(dataframes, ignore_index = True)

    df['deleted'] = True;
    transaction_rows, _ = np.where(df == 'Domestic Transactions');
    last_row, _ = np.where(df == 'Reward Points Summary');
    transaction_rows = np.concatenate([transaction_rows, last_row])
    for j in range(0, len(transaction_rows) - 1):
        handle_transaction(df, transaction_rows[j] ,transaction_rows[j + 1])
    df = df[df['deleted'] != True]
    del df['deleted']
    new_record = pd.DataFrame([['Date', 'to', 'amount', 'card']])
    df = pd.concat([new_record, df], ignore_index=True)

    #df = df.drop([i for i in range(0, 13)])
    return df


def main(args):
    for file_name in os.listdir(args.in_dir):
        root, ext = os.path.splitext(file_name)
        if ext.lower() != '.pdf':
            continue
        pdf_path = os.path.join(args.in_dir, file_name)
        print(f'Processing: {pdf_path}')
        df = extract_df(pdf_path)
        excel_name = root + '.xlsx'
        excel_path = os.path.join(args.out_dir, excel_name)
        df.to_excel(excel_path, index = False, header = False)
        print(f'Processed : {excel_path}')
        

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--in-dir', type=str, required=True, help='directory to read statement PDFs from.')
    parser.add_argument('--out-dir', type=str, required=True, help='directory to store statement XLSX to.')
    #parser.add_argument('--password', type=str, default=None, help='password for the statement PDF.')
    args = parser.parse_args()

    main(args)

