import math

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# data = {'Date': {0: '21/02/2021'}, 'time': {0: '66/66'}, 'Index': {0: 'NIFTY'},
#         'Put Write (Teji Wale)': {0: 43610.0}, 'Call Write ( Mandi Wale)': {0: 153896.0}, 'Diff (P-C)': {0: -110286.0},
#         'Diff (C-P)': {0: 110286.0}, 'P/C (Teji Jada)': {0: 0.2833731870873837},
#         'c/p (Mandi Jyada)': {0: 3.528915386379271}}
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']


def prepare_header(row):
    row = str(row)
    list = [{
        'range': 'A' + row,
        'values': [['Date']],
    }, {
        'range': 'B' + row,
        'values': [['Time']],
    }, {
        'range': 'C' + row,
        'values': [['Index']],
    }, {
        'range': 'D' + row,
        'values': [['Put Write (Teji Wale)']],
    }, {
        'range': 'E' + row,
        'values': [['Call Write ( Mandi Wale)']],
    }, {
        'range': 'F' + row,
        'values': [['Diff (P-C)']],
    }, {
        'range': 'G' + row,
        'values': [['Diff (C-P)']],
    }, {
        'range': 'H' + row,
        'values': [['P/C (Teji Jada)']],
    }, {
        'range': 'I' + row,
        'values': [['c/p (Mandi Jyada)']],
    }, {
        'range': 'J' + row,
        'values': [['Put Total OI']],
    }, {
        'range': 'K' + row,
        'values': [['Call Total OI']],
    }, {
        'range': 'L' + row,
        'values': [['P-C OI']],
    }, {
        'range': 'M' + row,
        'values': [['C-P OI']],
    }, {
        'range': 'N' + row,
        'values': [['P/C OI']],
    }, {
        'range': 'O' + row,
        'values': [['c/p OI']],
    }]

    return list


def prepare_data(row, data):
    if row > 1:
        row = str(row)
        a_val = data['Date'][0]
        b_val = data['time'][0]
        c_val = data['Index'][0]
        d_val = data['Put Write (Teji Wale)'][0]
        e_val = data['Call Write ( Mandi Wale)'][0]
        f_val = data['Diff (P-C)'][0]
        g_val = data['Diff (C-P)'][0]
        h_val = data['P/C (Teji Jada)'][0]
        i_val = data['c/p (Mandi Jyada)'][0]
        j_val = data['Put Total OI'][0]
        k_val = data['Call Total OI'][0]
        l_val = data['P-C OI'][0]
        m_val = data['C-P OI'][0]
        n_val = data['P/C OI'][0]
        o_val = data['c/p OI'][0]
        list = [{
            'range': 'A' + row,
            'values': [[a_val]],
        }, {
            'range': 'B' + row,
            'values': [[str(b_val)]],
        }, {
            'range': 'C' + row,
            'values': [[c_val]],
        }, {
            'range': 'D' + row,
            'values': [[d_val]],
        }, {
            'range': 'E' + row,
            'values': [[e_val]],
        }, {
            'range': 'F' + row,
            'values': [[f_val if math.isnan(f_val) else f_val]],
        }, {
            'range': 'G' + row,
            'values': [[g_val if math.isnan(g_val) else g_val]],
        }, {
            'range': 'H' + row,
            'values': [[h_val if math.isnan(h_val) else h_val]],
        }, {
            'range': 'I' + row,
            'values': [[i_val if math.isnan(i_val) else i_val]],
        }, {
            'range': 'J' + row,
            'values': [[j_val]],
        }, {
            'range': 'K' + row,
            'values': [[k_val]],
        }, {
            'range': 'L' + row,
            'values': [[l_val]],
        }, {
            'range': 'M' + row,
            'values': [[m_val]],
        }, {
            'range': 'N' + row,
            'values': [[n_val]],
        }, {
            'range': 'O' + row,
            'values': [[o_val]],
        }]

    return list


def main(row_num, data):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        'service_acc.json', scope)
    gc = gspread.authorize(credentials)
    wb = gc.open('FnO')
    if row_num <= len(wb.sheet1.get_all_values()):
        wb.sheet1.clear()
    wb.sheet1.batch_update(prepare_header(1))
    wb.sheet1.batch_update(prepare_data(row_num, data))
    return len(wb.sheet1.get_all_values())
