import pandas as pd
import math
from itertools import groupby
import logging
from datetime import datetime
import os


def converter(source_file):
    # file = r'data/Text (1).xlsx'
    file = source_file
    df = pd.read_excel(file)
    df.dropna(how='all')
    df.columns = df.columns.str.strip()
    # df.columns = df.columns.str.replace(' ', '')
    # df = df.iloc[N: , :]
    records = df.to_dict('records')
    data = []
    x_records = [x for x in records if x['Nợ - có'] not in ['2-2', '3-3', '4-4', '5-5', '6-6', '7-7', '8-8', '9-9', '2-3', '3-2', '4-2', '4-3','4-5', '5-4'] and x['GL No']]
    for key, values in groupby(x_records, key=lambda x: x['GL No']):
        list_vals_x = list(values)
        if isinstance(list_vals_x[0]['Nợ - có'], str) and '-' in list_vals_x[0]['Nợ - có']:
            check = list_vals_x[0]['Nợ - có'].split('-')
            dr = int(check[0])
            cr = int(check[1])
            if dr <= cr:
                # find item in list_vals_x has DrVND and DrVND not nan
                check_row = next((x for x in list_vals_x if x['DrVND'] and not math.isnan(x['DrVND'])), None)
                if check_row:
                    for item in list(filter(lambda x: x['CrVND'], list_vals_x)):
                        row = {}
                        row['GL Date'] = check_row['GL Date'].strftime('%m-%d-%Y') if check_row['GL Date'] and not isinstance(check_row['GL Date'], float) else ''
                        row['Posted Date'] = check_row['Posted Date'].strftime('%m-%d-%Y') if check_row['Posted Date'] and not isinstance(check_row['Posted Date'], float) else ''
                        row['GL No'] = check_row['GL No']
                        row['Description'] = check_row['Description']
                        row['Ledger-DrVND'] = check_row['Ledger']
                        row['AcName-DrVND'] = check_row['AcName']
                        row['Site'] = check_row['Site']
                        row['Nợ'] = check_row['Nợ']
                        row['Nợ - có'] = check_row['Nợ - có']
                        row['DrVND'] = item['CrVND']
                        row['DrUSD'] = item['CrUSD'] if item['CrUSD'] else 0
                        row['Amount-DrVND'] = item['Amount']
                        row['Amount-CrVND'] = item['Amount']
                        row['Ledger-CrVND'] = item['Ledger']
                        row['AcName-CrVND'] = item['AcName']
                        row['CrVND'] = item['CrVND']
                        row['CrUSD'] = item['CrUSD'] if item['CrUSD'] else 0
                        row['Có'] = item['Có']
                        data.append(row)
            else:
                check_row = next((x for x in list_vals_x if x['CrVND'] and not math.isnan(x['CrVND'])), None)
                if check_row:
                    for item in list(filter(lambda x: x['DrVND'], list_vals_x)):
                        row = {}
                        row['GL Date'] = check_row['GL Date'].strftime('%m-%d-%Y') if check_row['GL Date'] and not isinstance(check_row['GL Date'], float) else ''
                        row['Posted Date'] = check_row['Posted Date'].strftime('%m-%d-%Y') if check_row['Posted Date'] and not isinstance(check_row['Posted Date'], float) else ''
                        row['GL No'] = check_row['GL No']
                        row['Description'] = check_row['Description']
                        row['Ledger-CrVND'] = check_row['Ledger']
                        row['AcName-CrVND'] = check_row['AcName']
                        row['Site'] = check_row['Site']
                        row['Nợ'] = check_row['Nợ']
                        row['Nợ - có'] = check_row['Nợ - có']
                        row['CrVND'] = item['DrVND']
                        row['CrUSD'] = item['DrUSD'] if item['DrUSD'] else 0
                        row['Amount-CrVND'] = item['Amount']
                        row['Amount-DrVND'] = item['Amount']
                        row['Ledger-DrVND'] = item['Ledger']
                        row['AcName-DrVND'] = item['AcName']
                        row['DrVND'] = item['DrVND']
                        row['DrUSD'] = item['DrUSD'] if item['DrUSD'] else 0
                        row['Nợ'] = item['Nợ']
                        data.append(row)

    y_records = [y for y in records if y['Nợ - có'] in ['2-2', '3-3', '4-4', '5-5', '6-6', '7-7', '8-8', '9-9', '3-2', '2-3', '4-2', '4-3'] and y['GL No']]
    for key, values in groupby(y_records, key=lambda x: x['GL No']):
        list_vals_y = list(values)
        if isinstance(list_vals_y[0]['Nợ - có'], str) and '-' in list_vals_y[0]['Nợ - có']:
            check = list_vals_y[0]['Nợ - có'].split('-')
            dr = int(check[0])
            cr = int(check[1])

            list_vals_y_usd = list_vals_y

            list_vals_y_dr = sorted(list_vals_y, key=lambda d: d['DrVND'], reverse=True)
            check_max_dr = next((x for x in list_vals_y_dr if x['DrVND']), None)

            list_vals_y_cr = sorted(list_vals_y, key=lambda d: d['CrVND'], reverse=True)
            check_max_cr = next((y for y in list_vals_y_cr if y['CrVND']), None)

            list_cr = [x for x in list_vals_y if x['CrVND'] and x['CrVND'] != check_max_cr['CrVND']]
            list_dr = [y for y in list_vals_y if y['DrVND'] and y['DrVND'] != check_max_dr['DrVND']]

            check_usd = next((x for x in list_vals_y_usd if x['DrUSD'] or x['CrUSD']), None)
            list_vals_y_dr_usd = False
            list_vals_y_cr_usd = False
            check_max_dr_usd = False
            check_max_cr_usd = False
            list_cr_usd = False
            list_dr_usd = False
            if check_usd:
                list_vals_y_dr_usd = sorted(list_vals_y_usd, key=lambda d: d['DrUSD'], reverse=True)
                check_max_dr_usd = next((x for x in list_vals_y_dr_usd if x['DrUSD']), None)
                list_vals_y_cr_usd = sorted(list_vals_y_usd, key=lambda d: d['CrUSD'], reverse=True)
                check_max_cr_usd = next((y for y in list_vals_y_cr_usd if y['CrUSD']), None)
                list_cr_usd = [x for x in list_vals_y_cr_usd if x['CrUSD'] and x['CrUSD'] != check_max_cr_usd['CrUSD']]
                list_dr_usd = [y for y in list_vals_y_dr_usd if y['DrUSD'] and y['DrUSD'] != check_max_dr_usd['DrUSD']]

            if dr == cr:
                if check_max_dr and check_max_cr:
                    if check_max_cr['CrVND'] == check_max_dr['DrVND']:
                        list_cr = [x for x in list_vals_y if x['CrVND']]
                        list_dr = [y for y in list_vals_y if y['DrVND']]
                        for cr in list_cr:
                            check_dr = next((x for x in list_dr if x['DrVND'] == cr['CrVND']))
                            if check_dr:
                                row = {}
                                row['GL Date'] = cr['GL Date'].strftime('%m-%d-%Y') if cr['GL Date'] and not isinstance(cr['GL Date'], float) else ''
                                row['Posted Date'] = cr['Posted Date'].strftime('%m-%d-%Y') if cr['Posted Date'] and not isinstance(cr['Posted Date'], float) else ''
                                row['GL No'] = cr['GL No']
                                row['Description'] = cr['Description']
                                row['Ledger-CrVND'] = cr['Ledger']
                                row['AcName-CrVND'] = cr['AcName']
                                row['Site'] = cr['Site']
                                row['Nợ'] = cr['Nợ']
                                row['Nợ - có'] = cr['Nợ - có']
                                row['CrVND'] = cr['CrVND']
                                row['CrUSD'] = cr['CrUSD'] if cr['CrUSD'] else 0
                                row['Amount-CrVND'] = cr['Amount']
                                row['Amount-DrVND'] = check_dr['Amount']
                                row['Ledger-DrVND'] = check_dr['Ledger']
                                row['AcName-DrVND'] = check_dr['AcName']
                                row['DrVND'] = check_dr['DrVND']
                                row['DrUSD'] = check_dr['DrUSD'] if check_dr['DrUSD'] else 0
                                data.append(row)
                    else:
                        process(data, check_max_dr, check_max_cr, list_cr, list_dr, check_max_dr_usd , check_max_cr_usd , list_cr_usd , list_dr_usd )
            else:
                process(data, check_max_dr, check_max_cr, list_cr, list_dr, check_max_dr_usd , check_max_cr_usd , list_cr_usd , list_dr_usd)

    columns = ['GL Date', 'Posted Date', 'GL No', 'Description', 'Amount-DrVND', 'Amount-CrVND', 'Ledger-DrVND', 'Ledger-CrVND',
               'AcName-DrVND', 'AcName-CrVND', 'DrVND', 'CrVND', 'DrUSD', 'CrUSD', 'Site', 'Nợ', 'Có', 'Nợ - có']
    dff = pd.DataFrame(data, columns=columns)
    dest_file = os.path.dirname(os.path.abspath(
        source_file)) + '\\' + 'converted-' + os.path.basename(source_file)
    dff.to_excel(dest_file, sheet_name='Converted', index=False)

def process(data, check_max_dr, check_max_cr, list_cr, list_dr, check_max_dr_usd , check_max_cr_usd , list_cr_usd , list_dr_usd):
    row_cr = {}
    row_cr['GL Date'] = check_max_cr['GL Date'].strftime('%m-%d-%Y') if check_max_cr['GL Date'] and not isinstance(check_max_cr['GL Date'], float) else ''
    row_cr['Posted Date'] = check_max_cr['Posted Date'].strftime('%m-%d-%Y') if check_max_cr['Posted Date'] and not isinstance(check_max_cr['Posted Date'], float) else ''
    row_cr['GL No'] = check_max_cr['GL No']
    row_cr['Site'] = check_max_cr['Site']
    row_cr['Nợ - có'] = check_max_cr['Nợ - có']
    row_cr['Description'] = check_max_cr['Description']
    row_cr['Ledger-CrVND'] = check_max_cr['Ledger']
    row_cr['AcName-CrVND'] = check_max_cr['AcName']
    row_cr['CrVND'] = check_max_cr['CrVND'] - sum(x['DrVND'] for x in list_dr)
    row_cr['CrUSD'] = check_max_cr_usd['CrUSD'] - sum(x['DrUSD'] for x in list_dr_usd) if check_max_cr_usd and list_dr_usd else 0
                # get from row_dr
    row_cr['Ledger-DrVND'] = check_max_dr['Ledger']
    row_cr['AcName-DrVND'] = check_max_dr['AcName']
    row_cr['DrVND'] = check_max_dr['DrVND'] - sum(x['CrVND'] for x in list_cr)
    row_cr['DrUSD'] = check_max_dr_usd['DrUSD'] - sum(x['CrUSD'] for x in list_cr_usd) if check_max_dr_usd and list_cr_usd else 0
    row_cr['Amount-DrVND'] = row_cr['DrVND']
    row_cr['Amount-CrVND'] = row_cr['CrVND']
    row_cr['Nợ'] = ''
    row_cr['Có'] = ''
    data.append(row_cr)
    for item in list_cr:
        create_cr_row(data, check_max_dr, item)
    for item in list_dr:
        create_dr_row(data, check_max_cr, item)


def create_dr_row(data, check_max_cr, check_other_in_dr):
    row_vat = {}
    row_vat['GL Date'] = check_max_cr['GL Date'].strftime('%m-%d-%Y') if check_max_cr['GL Date'] and not isinstance(check_max_cr['GL Date'], float) else ''
    row_vat['Posted Date'] = check_max_cr['Posted Date'].strftime('%m-%d-%Y') if check_max_cr['Posted Date'] and not isinstance(check_max_cr['Posted Date'], float) else ''
    row_vat['GL No'] = check_max_cr['GL No']
    row_vat['Site'] = check_max_cr['Site']
    row_vat['Nợ - có'] = check_max_cr['Nợ - có']
    row_vat['Description'] = check_max_cr['Description']
    row_vat['Ledger-CrVND'] = check_max_cr['Ledger']
    row_vat['AcName-CrVND'] = check_max_cr['AcName']

    row_vat['Ledger-DrVND'] = check_other_in_dr['Ledger']
    row_vat['AcName-DrVND'] = check_other_in_dr['AcName']
    row_vat['DrVND'] = check_other_in_dr['DrVND']
    row_vat['DrUSD'] = check_other_in_dr['DrUSD']
    row_vat['CrVND'] = check_other_in_dr['DrVND']
    row_vat['CrUSD'] = check_other_in_dr['DrUSD']
    row_vat['Amount-DrVND'] = check_other_in_dr['Amount']
    row_vat['Amount-CrVND'] = check_other_in_dr['Amount']
    row_vat['Nợ'] = ''
    row_vat['Có'] = ''
    data.append(row_vat)


def create_cr_row(data, check_max_dr, check_other_in_cr):
    row_vat = {}
    row_vat['GL Date'] = check_max_dr['GL Date'].strftime('%m-%d-%Y') if check_max_dr['GL Date'] and not isinstance(check_max_dr['GL Date'], float) else ''
    row_vat['Posted Date'] = check_max_dr['Posted Date'].strftime('%m-%d-%Y') if check_max_dr['Posted Date'] and not isinstance(check_max_dr['Posted Date'], float) else ''
    row_vat['GL No'] = check_max_dr['GL No']
    row_vat['Site'] = check_max_dr['Site']
    row_vat['Nợ - có'] = check_max_dr['Nợ - có']
    row_vat['Description'] = check_max_dr['Description']
    row_vat['Ledger-DrVND'] = check_max_dr['Ledger']
    row_vat['AcName-DrVND'] = check_max_dr['AcName']

    row_vat['Ledger-CrVND'] = check_other_in_cr['Ledger']
    row_vat['AcName-CrVND'] = check_other_in_cr['AcName']
    row_vat['DrVND'] = check_other_in_cr['CrVND']
    row_vat['DrUSD'] = check_other_in_cr['CrUSD']
    row_vat['CrVND'] = check_other_in_cr['CrVND']
    row_vat['CrUSD'] = check_other_in_cr['CrUSD']
    row_vat['Amount-DrVND'] = check_other_in_cr['Amount']
    row_vat['Amount-CrVND'] = check_other_in_cr['Amount']
    row_vat['Nợ'] = ''
    row_vat['Có'] = ''
    data.append(row_vat)
