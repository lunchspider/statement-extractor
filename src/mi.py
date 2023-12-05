import os
import argparse
import shutil
from xlsxwriter import Workbook 
from pypdf import PdfReader

col_list = ([
    'file_name',
    'Order Number',
    'Order Date',
    'Invoice Date',
    'Invoice Number',
    'Seller Legal Name',
    'Seller GSTIN',
    'Buyer GSTIN',
    'Buyer Legal Name',
    'name',
    'hsn',
    'qty',
    'assesable_value',
    'igst',
    'cgst',
    'sgst',
    'total amount',
    'TCS',
])



def handle_file(file_name : str) -> list[dict[str, str]]:
    reader = PdfReader(file_name)
    page = reader.pages[0]
    text : str = page.extract_text()
    arr = text.split('\n')
    info = {'file_name': file_name}
    products : list[dict[str, str]] = []
    result = []
    for (index, i) in enumerate(arr):
        if 'Seller Legal Name' in i:
            name = i.replace('Seller Legal Name', '').strip()
            pos = name.index('CIN')
            info['Seller Legal Name'] = name[:pos].strip()
        if 'Buyer Legal Name' in i:
            name = i.replace('Buyer Legal Name', '').strip()
            index = name.index('Ship')
            info['Buyer Legal Name'] = name[:index].strip()
        if 'Order Number' in i:
            info['Order Number'] = i.split(' ')[-1].strip()
        if 'Order Date' in i:
            info['Order Date'] = i.split(' ')[-1].strip()
        if 'Invoice Number' in i:
            info['Invoice Number'] = i.split(' ')[-1].strip()
        if 'Invoice Date' in i:
            info['Invoice Date'] = i.split(' ')[-1].strip()
        if 'Buyer GSTIN' in i:
            info['Buyer GSTIN'] = i.split(' ')[2].strip()
        if 'Seller GSTIN' in i:
            info['Seller GSTIN'] = i.split(' ')[2].strip()
        if 'TCS' in i:
            info['TCS'] = i.split(' ')[1].strip()
        if 'Item Total' in i:
            pos = index
            while True:
                try:
                    name : str = arr[pos + 1].strip()
                    if 'shipping' in name or 'Total' in name:
                        break
                    l = []
                    if 'Nos' in arr[pos + 2]:
                        l = arr[pos+2].split(' ')
                        pos += 2
                    else:
                        l = arr[pos+3].split(' ')
                        name = name + ' ' + arr[pos + 2]
                        pos += 3
                    if (x := [index for index, st in enumerate(l) if len(st) == 12][0]) != 0:
                        name =  name + ' ' + ' '.join(l[:x])
                    index_of_nos = l.index('Nos')
                    hsn = l[index_of_nos - 3]
                    qty = l[index_of_nos - 1]
                    total_amount = l[index_of_nos + 13]
                    assesable_value = l[index_of_nos + 4]
                    igst = l[index_of_nos + 9]
                    cgst = l[index_of_nos + 10]
                    sgst = l[index_of_nos + 11]
                    products.append({ 
                        'name': name, 
                        'hsn': hsn, 
                        'qty': qty, 
                        'total amount': total_amount,
                        'assesable_value' : assesable_value,
                        'igst' : igst,
                        'cgst' : cgst,
                        'sgst' : sgst,
                    });
                except:
                    break
    for i in products:
        result.append({**info , **i})
    return result



def main(args):
    result : list[dict[str, str]]= []
    if not os.path.isdir(args.out_dir):
        os.makedirs(args.out_dir)
    for file_name in os.listdir(args.in_dir):
        root, ext = os.path.splitext(file_name)
        if ext.lower() != '.pdf':
            continue
        pdf_path = os.path.join(args.in_dir, file_name)
        print(f'Processing: {pdf_path}')

        try:
            info = handle_file(pdf_path)

            # file is possibly curropted!
            for info  in info:
                if sorted(info.keys()) != sorted(col_list):
                    raise SystemError('File curropeted')

                out_path = os.path.join(args.out_dir, f'{info["Buyer GSTIN"]} {info["Order Number"]} {info["Invoice Number"]}.pdf')
                shutil.copyfile(pdf_path, out_path)
                result.append(info)
        except Exception as ex:
            print(ex)
            print('file curropted!')
            if not os.path.isdir('./curropted'):
                os.makedirs('./curropted')
            out_path = os.path.join('./curropted/', file_name)
            shutil.copyfile(pdf_path, out_path)
        break

    wb = Workbook(args.out_file)
    ws=wb.add_worksheet("New Sheet")
    first_row = 0
    for header in col_list:
        col=col_list.index(header) # We are keeping order.
        ws.write(first_row,col,header) # We have written first row which is the header of worksheet also.
    row = 1
    for item in result:
        for _key,_value in item.items():
            col=col_list.index(_key)
            ws.write(row,col,_value)
        row+=1 #enter the next row
    wb.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--in-dir', type=str, required=True, help='directory to read statement PDFs from.')
    parser.add_argument('--out-file', type=str, required=True, help='file to store the output excel.')
    parser.add_argument('--out-dir', type=str, required=True, help='directory to write PDFs to.')
    #parser.add_argument('--password', type=str, default=None, help='password for the statement PDF.')
    args = parser.parse_args()

    main(args)



