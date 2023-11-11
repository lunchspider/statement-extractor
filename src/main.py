import os
import argparse
import pandas as pd
import shutil
from pypdf import PdfReader

def handle_file(file_name : str) -> dict[str, str]:
    reader = PdfReader(file_name)
    page = reader.pages[0]
    text = page.extract_text()
    arr = text.split('\n')
    info = {}
    for (index, i) in enumerate(arr):
        if 'Sold By' in i:
            x = i.replace('Sold By: ', '')
            x = x.replace(',', '')
            x = x.strip()
            info['Seller Name'] = x
        if 'Bill To' in i:
            buyer = arr[index + 1].strip()
            info['Buyer Name'] = buyer
        if 'Invoice Date:' in i:
            invoice_date = i.replace('Invoice Date:', '').strip()
            info['invoice_date'] = invoice_date
        if 'Order Date:' in i:
            order_date = i.replace('Order Date:', '').strip()
            info['order_date'] = order_date
        if 'GSTIN' in i:
            gstin = i.replace('GSTIN  -', '').strip()
            info['Seller GSTIN'] = gstin
        if 'Invoice Number' in i:
            invoice_number = i.replace('Invoice Number #', '').replace('Tax Invoice', '').strip()
            info['invoice_number']  = invoice_number
        if 'IMEI' in i:
            imei = i.split(' ')
            imei = ''.join(imei[len(imei) - 2: len(imei)])
            info['imei'] = imei
        if 'Grand Total' in i:
            i = arr[index - 1]
            s = i.split(' ')
            info['qty'] = s[1]
            info['total'] = s[-1]
            info['IGST'] = s[-2]
            info['taxable_value'] = s[-3]
        if 'Order ID:' in i:
            info['Buyer GSTIN'] = i.split(' ')[0]
            info['order_id'] = arr[index + 1].strip()
        if 'HSN' in i:
            s = i.replace('HSN/SAC: ', '').strip()
            hsn = ''
            for j in s:
                if not j.isdigit():
                    break
                hsn += j
            info['HSN'] = hsn
            s = s.replace(hsn, '') + arr[index + 1].strip()
            info['product_name'] = s
    return info


def main(args):
    result = []
    if not os.path.isdir(args.out_dir):
        os.makedirs(args.out_dir)
    for file_name in os.listdir(args.in_dir):
        root, ext = os.path.splitext(file_name)
        if ext.lower() != '.pdf':
            continue
        pdf_path = os.path.join(args.in_dir, file_name)
        print(f'Processing: {pdf_path}')
        info = handle_file(pdf_path)
        out_path = os.path.join(args.out_dir, f'{info["Buyer GSTIN"]}-{info["order_id"]}-{info["invoice_number"]}.pdf')
        shutil.copyfile(pdf_path, out_path)
        result.append(info)
    df = pd.DataFrame(result)
    df.to_excel(args.out_file, index = False)

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--in-dir', type=str, required=True, help='directory to read statement PDFs from.')
    parser.add_argument('--out-file', type=str, required=True, help='file to store the output excel.')
    parser.add_argument('--out-dir', type=str, required=True, help='directory to write PDFs to.')
    #parser.add_argument('--password', type=str, default=None, help='password for the statement PDF.')
    args = parser.parse_args()

    main(args)

