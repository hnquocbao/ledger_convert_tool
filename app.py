import pandas as pd
from itertools import groupby
import logging
from datetime import datetime
import os

def converter(source_file):
    # file = r'data/Text (1).xlsx'
    file = source_file
    df = pd.read_excel(file)
    df.dropna(how='all')
    # df = df.iloc[N: , :]
    records = df.to_dict('records')
    data = []
    x_records = [x for x in records if x['Nợ - có'] not in ['2-2', '3-2', '4-2', '4-3'] and x['GL No']]
    for key, values in groupby(x_records, key=lambda x: x['GL No']):
        list_vals_x = list(values)
        if isinstance(list_vals_x[0]['Nợ - có'], str) and '-' in list_vals_x[0]['Nợ - có']:
            check = list_vals_x[0]['Nợ - có'].split('-')
            dr = int(check[0])
            cr = int(check[1])
            if dr <= cr:
                check_row = next((x for x in list_vals_x if x['DrVND']), None)
                if check_row:
                    for item in list(filter(lambda x: x['CrVND'], list_vals_x)):
                        row = {}
                        row['GL Date'] = check_row['GL Date'].strftime(
                            '%m-%d-%Y')
                        row['Posted Date'] = check_row['Posted Date'].strftime(
                            '%m-%d-%Y')
                        row['GL No'] = check_row['GL No']
                        row['Description'] = check_row['Description']
                        row['Ledger-DrVND'] = check_row['Ledger']
                        row['AcName-DrVND'] = check_row['AcName']
                        row['Site'] = check_row['Site']
                        row['Nợ'] = check_row['Nợ']
                        row['Nợ - có'] = check_row['Nợ - có']
                        row['DrVND'] = item['CrVND']
                        row['Amount-DrVND'] = item['Amount']
                        row['Amount-CrVND'] = item['Amount']
                        row['Ledger-CrVND'] = item['Ledger']
                        row['AcName-CrVND'] = item['AcName']
                        row['CrVND'] = item['CrVND']
                        row['Có'] = item['Có']
                        data.append(row)
            else:
                check_row = next((x for x in list_vals_x if x['CrVND']), None)
                if check_row:
                    for item in list(filter(lambda x: x['DrVND'], list_vals_x)):
                        row = {}
                        row['GL Date'] = check_row['GL Date'].strftime(
                            '%m-%d-%Y')
                        row['Posted Date'] = check_row['Posted Date'].strftime(
                            '%m-%d-%Y')
                        row['GL No'] = check_row['GL No']
                        row['Description'] = check_row['Description']
                        row['Ledger-CrVND'] = check_row['Ledger']
                        row['AcName-CrVND'] = check_row['AcName']
                        row['Site'] = check_row['Site']
                        row['Nợ'] = check_row['Nợ']
                        row['Nợ - có'] = check_row['Nợ - có']
                        row['CrVND'] = item['DrVND']
                        row['Amount-CrVND'] = item['Amount']
                        row['Amount-DrVND'] = item['Amount']
                        row['Ledger-DrVND'] = item['Ledger']
                        row['AcName-DrVND'] = item['AcName']
                        row['DrVND'] = item['DrVND']
                        row['Nợ'] = item['Nợ']
                        data.append(row)

    y_records = [y for y in records if y['Nợ - có'] in ['2-2', '3-2', '2-3', '4-2'] and y['GL No']]
    for key, values in groupby(y_records, key=lambda x: x['GL No']):
        list_vals_y = list(values)
        if isinstance(list_vals_y[0]['Nợ - có'], str) and '-' in list_vals_y[0]['Nợ - có']:
            check = list_vals_y[0]['Nợ - có'].split('-')
            dr = int(check[0])
            cr = int(check[1])
            if dr == cr:
                list_vals_y_dr = sorted(list_vals_y, key=lambda d: d['DrVND'], reverse=True)
                check_max_dr = next((x for x in list_vals_y_dr if x['DrVND']), None)
                list_vals_y_cr = sorted(list_vals_y, key=lambda d: d['CrVND'], reverse=True)
                check_max_cr = next((y for y in list_vals_y_cr if y['CrVND']), None)
                if check_max_dr and check_max_cr:
                    if check_max_cr['CrVND'] > check_max_dr['DrVND']:
                        # run check_max_cr
                        check_vat_in_dr = next((x for x in list_vals_y if x['DrVND'] and 'VAT' in x['AcName'] or x['Ledger'] == 133111), None)
                        check_purchase_discount_in_cr = next((y for y in list_vals_y if y['CrVND'] and 'PURCHASE' in y['AcName'] or y['Ledger'] == 6327), None)
                        if check_vat_in_dr and check_purchase_discount_in_cr:
                            new_func(data, check_max_cr, check_vat_in_dr)
                            row_purchase_discount = {}
                            row_purchase_discount['GL Date'] = check_max_dr['GL Date'].strftime('%m-%d-%Y')
                            row_purchase_discount['Posted Date'] = check_max_dr['Posted Date'].strftime('%m-%d-%Y')
                            row_purchase_discount['GL No'] = check_max_dr['GL No']
                            row_purchase_discount['Site'] = check_max_dr['Site']
                            row_purchase_discount['Nợ - có'] = check_max_dr['Nợ - có']
                            row_purchase_discount['Description'] = check_max_dr['Description']
                            row_purchase_discount['Ledger-DrVND'] = check_max_dr['Ledger']
                            row_purchase_discount['AcName-DrVND'] = check_max_dr['AcName']

                            row_purchase_discount['Ledger-CrVND'] = check_purchase_discount_in_cr['Ledger']
                            row_purchase_discount['AcName-CrVND'] = check_purchase_discount_in_cr['AcName']
                            row_purchase_discount['DrVND'] = check_purchase_discount_in_cr['CrVND']
                            row_purchase_discount['CrVND'] = check_purchase_discount_in_cr['CrVND']
                            row_purchase_discount['Amount-DrVND'] = check_purchase_discount_in_cr['Amount']
                            row_purchase_discount['Amount-CrVND'] = check_purchase_discount_in_cr['Amount']

                            row_purchase_discount['Nợ'] = ''
                            row_purchase_discount['Có'] = ''
                            data.append(row_purchase_discount)

                            row_cr = {}
                            row_cr['GL Date'] = check_max_cr['GL Date'].strftime('%m-%d-%Y')
                            row_cr['Posted Date'] = check_max_cr['Posted Date'].strftime('%m-%d-%Y')
                            row_cr['GL No'] = check_max_cr['GL No']
                            row_cr['Site'] = check_max_cr['Site']

                            row_cr['Nợ - có'] = check_max_cr['Nợ - có']
                            row_cr['Description'] = check_max_cr['Description']

                            row_cr['Ledger-CrVND'] = check_max_cr['Ledger']
                            row_cr['AcName-CrVND'] = check_max_cr['AcName']
                            row_cr['CrVND'] = check_max_cr['CrVND'] - check_vat_in_dr['DrVND']
                            # get from row_dr
                            row_cr['Ledger-DrVND'] = check_purchase_discount_in_cr['Ledger']
                            row_cr['AcName-DrVND'] = check_purchase_discount_in_cr['AcName']
                            row_cr['DrVND'] = check_max_dr['DrVND'] - check_purchase_discount_in_cr['CrVND']

                            row_cr['Amount-DrVND'] = row_cr['DrVND']
                            row_cr['Amount-CrVND'] = row_cr['CrVND']

                            row_cr['Nợ'] = ''
                            row_cr['Có'] = ''
                            data.append(row_cr)
                    elif check_max_cr['CrVND'] < check_max_dr['DrVND']:
                    # run check_max_dr
                        check_vat_in_cr = next((x for x in list_vals_y if x['CrVND'] and 'VAT' in x['AcName'] or x['Ledger'] == 133111), None)
                        check_purchase_discount_in_dr = next((y for y in list_vals_y if y['DrVND'] and 'PURCHASE' in y['AcName'] or y['Ledger'] == 6327), None)
                        if check_vat_in_cr and check_purchase_discount_in_dr:
                            row_vat = {}
                            row_vat['GL Date'] = check_max_dr['GL Date'].strftime('%m-%d-%Y')
                            row_vat['Posted Date'] = check_max_dr['Posted Date'].strftime('%m-%d-%Y')
                            row_vat['GL No'] = check_max_dr['GL No']
                            row_vat['Site'] = check_max_dr['Site']

                            row_vat['Nợ - có'] = check_max_dr['Nợ - có']
                            row_vat['Description'] = check_max_dr['Description']

                            row_vat['Ledger-CrVND'] = check_max_dr['Ledger']
                            row_vat['AcName-CrVND'] = check_max_dr['AcName']
                            row_vat['Ledger-DrVND'] = check_vat_in_cr['Ledger']
                            row_vat['AcName-DrVND'] = check_vat_in_cr['AcName']
                            row_vat['DrVND'] = check_vat_in_cr['CrVND']
                            row_vat['CrVND'] = check_vat_in_cr['CrVND']
                            row_vat['Amount-DrVND'] = check_vat_in_cr['Amount']
                            row_vat['Amount-CrVND'] = check_vat_in_cr['Amount']

                            row_vat['Nợ'] = ''
                            row_vat['Có'] = ''
                            data.append(row_vat)

                            row_purchase_discount = {}
                            row_purchase_discount['GL Date'] = check_max_dr['GL Date'].strftime('%m-%d-%Y')
                            row_purchase_discount['Posted Date'] = check_max_dr['Posted Date'].strftime('%m-%d-%Y')
                            row_purchase_discount['GL No'] = check_max_dr['GL No']
                            row_purchase_discount['Site'] = check_max_dr['Site']

                            row_purchase_discount['Nợ - có'] = check_max_dr['Nợ - có']
                            row_purchase_discount['Description'] = check_max_dr['Description']
                            row_purchase_discount['Ledger-DrVND'] = check_max_dr['Ledger']
                            row_purchase_discount['AcName-DrVND'] = check_max_dr['AcName']

                            row_purchase_discount['Ledger-CrVND'] = check_purchase_discount_in_dr['Ledger']
                            row_purchase_discount['AcName-CrVND'] = check_purchase_discount_in_dr['AcName']
                            row_purchase_discount['DrVND'] = check_purchase_discount_in_dr['DrVND']
                            row_purchase_discount['CrVND'] = check_purchase_discount_in_dr['DrVND']
                            row_purchase_discount['Amount-DrVND'] = check_purchase_discount_in_dr['Amount']
                            row_purchase_discount['Amount-CrVND'] = check_purchase_discount_in_dr['Amount']

                            row_purchase_discount['Nợ'] = ''
                            row_purchase_discount['Có'] = ''
                            data.append(row_purchase_discount)

                            row_dr = {}
                            row_dr['GL Date'] = check_max_dr['GL Date'].strftime('%m-%d-%Y')
                            row_dr['Posted Date'] = check_max_dr['Posted Date'].strftime('%m-%d-%Y')
                            row_dr['GL No'] = check_max_dr['GL No']
                            row_dr['Site'] = check_max_dr['Site']
                            row_dr['Nợ - có'] = check_max_dr['Nợ - có']
                            row_dr['Description'] = check_max_dr['Description']
                            row_dr['Ledger-CrVND'] = check_max_dr['Ledger']
                            row_dr['AcName-CrVND'] = check_max_dr['AcName']
                            row_dr['CrVND'] = check_max_cr['CrVND'] - check_purchase_discount_in_dr['DrVND']
                            # get from row_dr
                            row_dr['Ledger-DrVND'] = check_purchase_discount_in_dr['Ledger']
                            row_dr['AcName-DrVND'] = check_purchase_discount_in_dr['AcName']
                            row_dr['DrVND'] = check_max_dr['DrVND'] - check_vat_in_cr['CrVND']

                            row_dr['Amount-DrVND'] = row_dr['DrVND']
                            row_dr['Amount-CrVND'] = row_dr['CrVND']

                            row_dr['Nợ'] = ''
                            row_dr['Có'] = ''
                            data.append(row_dr)
                    elif check_max_cr['CrVND'] == check_max_dr['DrVND']:
                        list_cr = [x for x in list_vals_y if x['CrVND']]
                        list_dr = [y for y in list_vals_y if y['DrVND']]
                        for cr in list_cr:
                            check_dr = next((x for x in list_dr if x['DrVND'] == cr['CrVND']))
                            if check_dr:
                                row = {}
                                row['GL Date'] = cr['GL Date'].strftime('%m-%d-%Y')
                                row['Posted Date'] = cr['Posted Date'].strftime('%m-%d-%Y')
                                row['GL No'] = cr['GL No']
                                row['Description'] = cr['Description']
                                row['Ledger-CrVND'] = cr['Ledger']
                                row['AcName-CrVND'] = cr['AcName']
                                row['Site'] = cr['Site']
                                row['Nợ'] = cr['Nợ']
                                row['Nợ - có'] = cr['Nợ - có']
                                row['CrVND'] = cr['CrVND']
                                row['Amount-CrVND'] = cr['Amount']
                                row['Amount-DrVND'] = check_dr['Amount']
                                row['Ledger-DrVND'] = check_dr['Ledger']
                                row['AcName-DrVND'] = check_dr['AcName']
                                row['DrVND'] = check_dr['DrVND']
                                data.append(row)
            elif dr > cr:
                list_vals_y_dr = sorted(list_vals_y, key=lambda d: d['DrVND'], reverse=True)
                check_max_dr = next((x for x in list_vals_y_dr if x['DrVND']), None)
                list_vals_y_cr = sorted(list_vals_y, key=lambda d: d['CrVND'], reverse=True)
                check_max_cr = next((y for y in list_vals_y_cr if y['CrVND']), None)
                if check_max_dr and check_max_cr:
                    if check_max_cr['CrVND'] > check_max_dr['DrVND']:
                        check_vat_in_dr = next((x for x in list_vals_y if x['DrVND'] and 'VAT' in x['AcName'] or x['Ledger'] == 133111), None)
                        check_purchase_service_in_dr = next((y for y in list_vals_y if y['DrVND'] and 'PURCHASE - SERVICE CHARGE' in y['AcName'] or y['Ledger'] == 63210), None)
                        check_purchase_other_charge_in_dr = next((y for y in list_vals_y if y['DrVND'] and 'PURCHASE - OTHER CHARGE' in y['AcName'] or y['Ledger'] == 6329), None)
                        check_purchase_discount_in_cr = next((y for y in list_vals_y if y['CrVND'] and 'PURCHASE - DISCOUNT' in y['AcName'] or y['Ledger'] == 6327), None)
                        check_vat_in_dr_value = 0
                        check_purchase_service_in_dr_value = 0
                        check_purchase_other_charge_in_dr_value = 0
                        check_purchase_discount_in_cr_value = 0
                        if check_vat_in_dr:
                            check_vat_in_dr_value = check_vat_in_dr['DrVND']
                            new_func(data, check_max_cr, check_vat_in_dr)
                        if check_purchase_service_in_dr:
                            check_purchase_service_in_dr_value = check_purchase_service_in_dr['DrVND']
                            new_func(data, check_max_cr, check_purchase_service_in_dr)
                        if check_purchase_other_charge_in_dr:
                            check_purchase_other_charge_in_dr_value = check_purchase_other_charge_in_dr['DrVND']
                            new_func(data, check_max_cr, check_purchase_other_charge_in_dr)
                        if check_purchase_discount_in_cr:
                            check_purchase_discount_in_cr_value = check_purchase_discount_in_cr['CrVND']
                            row_purchase_discount = {}
                            row_purchase_discount['GL Date'] = check_max_cr['GL Date'].strftime('%m-%d-%Y')
                            row_purchase_discount['Posted Date'] = check_max_cr['Posted Date'].strftime('%m-%d-%Y')
                            row_purchase_discount['GL No'] = check_max_cr['GL No']
                            row_purchase_discount['Site'] = check_max_cr['Site']
                            row_purchase_discount['Nợ - có'] = check_max_cr['Nợ - có']
                            row_purchase_discount['Description'] = check_max_cr['Description']

                            row_purchase_discount['Ledger-CrVND'] = check_purchase_discount_in_cr['Ledger']
                            row_purchase_discount['AcName-CrVND'] = check_purchase_discount_in_cr['AcName']

                            row_purchase_discount['Ledger-DrVND'] = check_max_cr['Ledger']
                            row_purchase_discount['AcName-DrVND'] = check_max_cr['AcName']
                            row_purchase_discount['DrVND'] = check_purchase_discount_in_cr['CrVND']
                            row_purchase_discount['CrVND'] = check_purchase_discount_in_cr['CrVND']
                            row_purchase_discount['Amount-DrVND'] = check_purchase_discount_in_cr['Amount']
                            row_purchase_discount['Amount-CrVND'] = check_purchase_discount_in_cr['Amount']
                            row_purchase_discount['Nợ'] = ''
                            row_purchase_discount['Có'] = ''
                            data.append(row_purchase_discount)

                        row_cr = {}
                        row_cr['GL Date'] = check_max_cr['GL Date'].strftime('%m-%d-%Y')
                        row_cr['Posted Date'] = check_max_cr['Posted Date'].strftime('%m-%d-%Y')
                        row_cr['GL No'] = check_max_cr['GL No']
                        row_cr['Site'] = check_max_cr['Site']
                        row_cr['Nợ - có'] = check_max_cr['Nợ - có']
                        row_cr['Description'] = check_max_cr['Description']
                        row_cr['Ledger-CrVND'] = check_max_cr['Ledger']
                        row_cr['AcName-CrVND'] = check_max_cr['AcName']
                        row_cr['CrVND'] = check_max_cr['CrVND'] - check_vat_in_dr_value - check_purchase_service_in_dr_value - check_purchase_other_charge_in_dr_value
                        # get from row_dr
                        row_cr['Ledger-DrVND'] = check_max_dr['Ledger']
                        row_cr['AcName-DrVND'] = check_max_dr['AcName']
                        row_cr['DrVND'] = check_max_dr['DrVND'] - check_purchase_discount_in_cr_value
                        row_cr['Amount-DrVND'] = row_cr['DrVND']
                        row_cr['Amount-CrVND'] = row_cr['CrVND']
                        row_cr['Nợ'] = ''
                        row_cr['Có'] = ''
                        data.append(row_cr)       
            elif dr < cr:
                list_vals_y_dr = sorted(list_vals_y, key=lambda d: d['DrVND'], reverse=True)
                check_max_dr = next((x for x in list_vals_y_dr if x['DrVND']), None)
                list_vals_y_cr = sorted(list_vals_y, key=lambda d: d['CrVND'], reverse=True)
                check_max_cr = next((y for y in list_vals_y_cr if y['CrVND']), None)
                if check_max_dr and check_max_cr:
                    if check_max_cr['CrVND'] < check_max_dr['DrVND']:
                        list_cr = [x for x in list_vals_y if x['CrVND'] and x['CrVND'] != check_max_cr['CrVND']]
                        list_dr = [y for y in list_vals_y if y['DrVND']]
                        check_vat_in_dr = next((x for x in list_vals_y if x['DrVND'] and 'VAT' in x['AcName'] or x['Ledger'] == 133111), None)
                        for item in list_cr:
                            row = {}
                            row['GL Date'] = check_max_dr['GL Date'].strftime('%m-%d-%Y')
                            row['Posted Date'] = check_max_dr['Posted Date'].strftime('%m-%d-%Y')
                            row['GL No'] = check_max_dr['GL No']
                            row['Description'] = check_max_dr['Description']
                            row['Ledger-DrVND'] = check_max_dr['Ledger']
                            row['AcName-DrVND'] = check_max_dr['AcName']
                            row['Site'] = check_max_dr['Site']
                            row['Nợ'] = check_max_dr['Nợ']
                            row['Nợ - có'] = check_max_dr['Nợ - có']
                            row['CrVND'] = item['CrVND']
                            row['DrVND'] = item['CrVND']
                            row['Amount-CrVND'] = item['Amount']
                            row['Amount-DrVND'] = item['Amount']
                            row['Ledger-CrVND'] = item['Ledger']
                            row['AcName-CrVND'] = item['AcName']
                            
                            data.append(row)

                        row_vat = {}
                        row_vat['GL Date'] = check_max_dr['GL Date'].strftime('%m-%d-%Y')
                        row_vat['Posted Date'] = check_max_dr['Posted Date'].strftime('%m-%d-%Y')
                        row_vat['GL No'] = check_max_dr['GL No']
                        row_vat['Site'] = check_max_dr['Site']

                        row_vat['Nợ - có'] = check_max_dr['Nợ - có']
                        row_vat['Description'] = check_max_dr['Description']

                        row_vat['Ledger-CrVND'] = check_max_dr['Ledger']
                        row_vat['AcName-CrVND'] = check_max_dr['AcName']
                        row_vat['Ledger-DrVND'] = check_vat_in_dr['Ledger']
                        row_vat['AcName-DrVND'] = check_vat_in_dr['AcName']
                        row_vat['DrVND'] = check_vat_in_dr['DrVND']
                        row_vat['CrVND'] = check_vat_in_dr['DrVND']
                        row_vat['Amount-DrVND'] = check_vat_in_dr['Amount']
                        row_vat['Amount-CrVND'] = check_vat_in_dr['Amount']

                        row_vat['Nợ'] = ''
                        row_vat['Có'] = ''
                        data.append(row_vat)
                        
                        row_dr = {}
                        row_dr['GL Date'] = check_max_dr['GL Date'].strftime('%m-%d-%Y')
                        row_dr['Posted Date'] = check_max_dr['Posted Date'].strftime('%m-%d-%Y')
                        row_dr['GL No'] = check_max_dr['GL No']
                        row_dr['Site'] = check_max_dr['Site']
                        row_dr['Nợ - có'] = check_max_dr['Nợ - có']
                        row_dr['Description'] = check_max_dr['Description']
                        row_dr['Ledger-CrVND'] = check_max_cr['Ledger']
                        row_dr['AcName-CrVND'] = check_max_cr['AcName']
                        row_dr['CrVND'] = check_max_cr['CrVND'] - check_vat_in_dr['DrVND']
                            # get from row_dr
                        row_dr['Ledger-DrVND'] = check_max_dr['Ledger']
                        row_dr['AcName-DrVND'] = check_max_dr['AcName']
                        row_dr['DrVND'] = check_max_dr['DrVND'] - sum([x['CrVND'] for x in list_cr])

                        row_dr['Amount-DrVND'] = row_dr['DrVND']
                        row_dr['Amount-CrVND'] = row_dr['CrVND']

                        row_dr['Nợ'] = ''
                        row_dr['Có'] = ''
                        data.append(row_dr)
                         

    columns = ['GL Date', 'Posted Date', 'GL No', 'Description', 'Amount-DrVND', 'Amount-CrVND', 'Ledger-DrVND', 'Ledger-CrVND',
               'AcName-DrVND', 'AcName-CrVND', 'DrVND', 'CrVND', 'DrUSD', 'CrUSD', 'Site', 'Nợ', 'Có', 'Nợ - có']
    dff = pd.DataFrame(data, columns=columns)
    dest_file = os.path.dirname(os.path.abspath(source_file)) + '\\' + 'converted-' + os.path.basename(source_file)
    dff.to_excel(dest_file, sheet_name='Converted', index=False)

def new_func(data, check_max_cr, check_other_in_dr):
    row_vat = {}
    row_vat['GL Date'] = check_max_cr['GL Date'].strftime('%m-%d-%Y')
    row_vat['Posted Date'] = check_max_cr['Posted Date'].strftime('%m-%d-%Y')
    row_vat['GL No'] = check_max_cr['GL No']
    row_vat['Site'] = check_max_cr['Site']
    row_vat['Nợ - có'] = check_max_cr['Nợ - có']
    row_vat['Description'] = check_max_cr['Description']
    row_vat['Ledger-CrVND'] = check_max_cr['Ledger']
    row_vat['AcName-CrVND'] = check_max_cr['AcName']

    row_vat['Ledger-DrVND'] = check_other_in_dr['Ledger']
    row_vat['AcName-DrVND'] = check_other_in_dr['AcName']
    row_vat['DrVND'] = check_other_in_dr['DrVND']
    row_vat['CrVND'] = check_other_in_dr['DrVND']
    row_vat['Amount-DrVND'] = check_other_in_dr['Amount']
    row_vat['Amount-CrVND'] = check_other_in_dr['Amount']
    row_vat['Nợ'] = ''
    row_vat['Có'] = ''
    data.append(row_vat)

