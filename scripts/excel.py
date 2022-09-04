import xlsxwriter
import pandas as pd
import logging
import os
import datetime
from win32com import client
import win32api

_logger = logging.getLogger("excel")

class Format:
    def __init__(self, workbook, cfg):
        self.heading = workbook.add_format(cfg['format_heading'])
        self.heading_orient = workbook.add_format(cfg['format_heading'])
        self.heading_orient.set_rotation(90)
        self.heading_orient.set_align('center')
        self.heading_center = workbook.add_format({**cfg['format_heading'], 'align': 'center'})
        self.heading_right = workbook.add_format({**cfg['format_heading'], 'align': 'right'})
        self.heading_left = workbook.add_format({**cfg['format_heading'], 'align': 'left'})
        self.center = workbook.add_format({'align': 'center'})
        self.right = workbook.add_format({'align': 'right'})
        self.even = workbook.add_format(cfg['format_row_even'])
        self.valid = workbook.add_format(cfg['format_number_valid'])
        self.invalid = workbook.add_format(cfg['format_number_invalid'])
    
def export(cfg, filename, data):
    workbook = xlsxwriter.Workbook(filename)
    fmt = Format(workbook, cfg)
    
    _logger.info('Write raw data to excel')
    write_raw_data(workbook, data, cfg, fmt)
    
    #_logger.info('Write controls data to excel')
    #write_controls_data(workbook, data, cfg, fmt)
   
    _logger.info('Write all available controls data')
    write_controls_full(workbook, data, cfg, fmt)
    
    _logger.info('Write patients data')
    write_patients_data(workbook, data, cfg, fmt)
   
    workbook.close()
    exceltopdf(filename)


def write_raw_data(workbook, data, cfg, fmt):
    ws_raw_data = workbook.add_worksheet('Rohdaten')
    ws_raw_data.set_landscape()
    write_maxtrix(0, 0, data['raw_data'], ws_raw_data, format_header=fmt.heading)


def write_controls_data(workbook, data, cfg, fmt):
    aoi = cfg['aminos_score']
    ws_controls = workbook.add_worksheet('Kontrollen')    
    ws_controls.set_landscape()
    ws_controls.set_header('&L&A' + '&CMessergebnisse des Aminosäure-Screenings' + '&RSeite &P von &N')
    ws_controls.set_footer('&RDatum: &D, &T')
    
    fmt_border = workbook.add_format({'bottom': 1, 'top': 1})
    fmt_border_left = workbook.add_format({'bottom': 1, 'top': 1, 'left': 1})
    fmt_border_right = workbook.add_format({'bottom': 1, 'top': 1, 'right': 1})
    
    ws_controls.write(0, 0, F"Bevorzugte Kontrolle: {data['selected_control']['name']}", fmt.heading)
    row_idx = 2
    ws_controls.write(row_idx, 0, 'Ranking', fmt.heading_orient)
    ws_controls.write(row_idx, 1, 'Kontrolle', fmt.heading_orient)
    ws_controls.write(row_idx, 2, 'Gültige AS', fmt.heading_orient)
    ws_controls.set_column('D:D', None, None, {'hidden': True})
    ws_controls.write(row_idx, 3, 'Score', fmt.heading_orient)
    ws_controls.write_row(row_idx, 5, data['controls'][aoi].columns.values.tolist(), fmt.heading_orient)
    ws_controls.set_column('A:Y', 4.3)
    ws_controls.set_column('B:C', 7.5)  # Name der Kontrolle
    ws_controls.set_column('C:C', 5.0)  # Gültigkeit
    ws_controls.set_column('E:E', 1.0)  # Abstand vor den aminos
    
    row_idx = 3
    first_invalid = None
    for rank, dat in enumerate(data['checked_controls']):
        ws_controls.write(row_idx, 0, rank+1, fmt.center)
        ws_controls.write(row_idx, 1, dat['name'], fmt.heading)
        ws_controls.write(row_idx, 2, f"{dat['coarse_score']}/20")
        ws_controls.write(row_idx, 3, f"{round(dat['fine_score']*100,1)}%", fmt.center)
        # highlight every second row to increase readability
        if rank%2 == 0:
            ws_controls.write_row(row_idx, 5, dat['raw_data'][aoi].to_numpy())
        else:
            ws_controls.write_row(row_idx, 5, dat['raw_data'][aoi].to_numpy(), fmt.even)
        
        res = dat['result'][aoi]
        # mark invalids 
        invalids = res[res == 'TOO_LOW'].index.to_list() + res[res == 'TOO_HIGH'].index.to_list()
        invalids = set(invalids)
        for inv in invalids:
            pos = res.index.get_loc(inv)
            ws_controls.conditional_format(row_idx, pos+5, row_idx, pos+5, {'type': 'no_errors', 'format': fmt.invalid}) 
            # mark heading as invalid for the selected control only
            if data['selected_control']['name'] == dat['name']:   #rank == 0 and 
                ws_controls.conditional_format(2, pos+5, 2, pos+5, {'type': 'no_errors', 'format': fmt.invalid}) 
        
        # add border to the selected with control
        if data['selected_control']['name'] == dat['name']:
            ws_controls.conditional_format(row_idx, 1, row_idx, 23, {'type': 'no_errors', 'format': fmt_border}) 
            ws_controls.conditional_format(row_idx, 0, row_idx, 0, {'type': 'no_errors', 'format': fmt_border_left})
            ws_controls.conditional_format(row_idx, 24, row_idx, 24, {'type': 'no_errors', 'format': fmt_border_right})
            
        row_idx += 1


def write_controls_full(workbook, data, cfg, fmt):
    ws_controls = workbook.add_worksheet('Kontrollen')
    ws_controls.set_header('&L&A' + '&CMessergebnisse des Aminosäure-Screenings' + '&RSeite &P von &N')
    ws_controls.set_footer('&RDatum: &D, &T')
    ws_controls.set_column('A:A', 24.0)
    ws_controls.set_column('B:B', 1.0)  # gap between the patient
    ws_controls.set_column('C:K', 8.0)  # all other columns

    # split up data into important aminos and secondary aminos
    aoi_names = cfg['aminos_score']
    missing_names = sorted(list(set(cfg['aminos_names'].values()) - set(aoi_names)))
    all_columns = list(data['controls'].columns.values)
    # re order control columns
    data['controls'] = data['controls'][all_columns[:2] + aoi_names + missing_names]

    # write first column
    # write score heading
    ws_controls.write(1, 0, 'Score', fmt.heading_right)
    # write all amino names as heading
    ws_controls.write_column(2, 0, data['controls'].columns.values.tolist()[2:], fmt.heading_right)
    len_aminos = len(data['controls'].columns.values.tolist()[2:])
    
    col_idx = 2
    for idx, dat in enumerate(data['checked_controls']):
        # write score
        _logger.info(f'Show coarse score for all aminos')
        ws_controls.write(1, col_idx+idx, f'{dat["coarse_score"]}/{len(cfg["aminos_score"])}', fmt.right)
        # write control name
        control_name = dat['name']
        ws_controls.write(0, col_idx+idx, control_name, fmt.heading_center)
        # re order control columns
        dat['raw_data'] = dat['raw_data'][all_columns[:2] + aoi_names + missing_names]
        # write results
        res = dat['raw_data'][2:].astype('float').round(2).to_numpy()
        ws_controls.write_column(2, col_idx+idx, res)

        # mark with gray all invalid aminos
        result = dat['result']
        invalids = result[result == 'TOO_LOW'].index.to_list() + result[result == 'TOO_HIGH'].index.to_list()
        invalids = set(invalids)
        amino_names = data['controls'].columns.values.tolist()[2:]
        _logger.info(f'Mark invalids: {invalids}')
        for inv in invalids:
            row_idx = amino_names.index(inv) + 2
            ws_controls.conditional_format(row_idx, 0, row_idx, 0, {'type': 'no_errors', 'format': fmt.invalid})
            ws_controls.conditional_format(row_idx, idx + 2, row_idx, idx + 2, {'type': 'no_errors', 'format': fmt.invalid})

    # add separation between aminos of interest and the rest of them
    fmt_border_bottom = workbook.add_format({'bottom': 1})
    ws_controls.conditional_format(28, 0, 28, len(data['checked_controls']) + 1, {'type': 'no_errors', 'format': fmt_border_bottom})
    ws_controls.conditional_format(1, 0, 1, len(data['checked_controls']) + 1, {'type': 'no_errors', 'format': fmt_border_bottom})

    # mark with gray every 3th, 4th row, 7th, 8th etc
    offset = 2
    col_len = len(data['checked_controls']) + 1
    for idx in list(range(len_aminos))[2::4]:
        ws_controls.conditional_format(idx+offset, 0, idx+offset, col_len, {'type': 'no_errors', 'format': fmt.even}) 
        ws_controls.conditional_format(idx+offset+1, 0, idx+offset+1, col_len, {'type': 'no_errors', 'format': fmt.even}) 

        
def write_patients_data(workbook, data, cfg, fmt):    
    ws_patients = workbook.add_worksheet('Patienten')
    ws_patients.set_landscape()
    ws_patients.set_header('&L&A' + '&CMessergebnisse des Aminosäure-Screenings' + '&RSeite &P von &N')
    ws_patients.set_footer('&RDatum: &D, &T')
    ws_patients.set_column('A:Y', 6.0)  # define all column width
    ws_patients.set_column('A:B', 5.2)
    ws_patients.set_column('C:C', 14.5)
    ws_patients.set_column('M:M', 14.5)
    ws_patients.set_column('H:H', 3.0)  # gap between the patient
    
    # format the patient data
    aoi = cfg['aminos_score']
    patients = data['data']
    pat_ref = data['patients_reference']
    fmt_map = pd.DataFrame().reindex_like(patients)
    fmt_map.iloc[:] = 0  # init with zeros
    fmt_map = fmt_map.astype('int')  # convert to integer
    
    # iterate by aminos
    for col in pat_ref.columns.values:
        val_min, val_max = pat_ref.loc[:, col]
        fmt_map[col][patients[col] < val_min] = 1  # mark as too low
        fmt_map[col][patients[col] > val_max] = 2  # mark as too high
    # mark aminos as invalid
    co_names = [d['name'] for d in data['checked_controls']]
    co_names_reduced = [n.replace('_1', '').replace('_2', '').replace('_3', '') for n in co_names]
    for co in set(co_names_reduced):
        co_idx_list = [i for i, n in enumerate(co_names_reduced) if n == co]
        invalids = set(data['checked_controls'][0]['result'].index)  # init with set of all
        for idx in co_idx_list:
            res = data['checked_controls'][idx]['result']
            invalids = invalids & set(res[(res == 'TOO_LOW') | (res == 'TOO_HIGH')].index)
        fmt_map[list(invalids)] = -1

    gap_rows = 36
    offset_col = 3
    idx_row = 0
    idx_col = offset_col
    second_part = 0
    
    # reformat array with data. Insert empty column
    amino_names = []
    amino_min = []
    amino_max = []
    add_empty_row = [4, 8, 12, 16, 20, 24, 28, 32]
   
    for idx, as_name in enumerate(aoi):                             
        if idx in add_empty_row:
            amino_names.append('')
            amino_min.append('')
            amino_max.append('')
        amino_names.append(as_name)
        amino_min.append(data['patients_reference'].loc[0, as_name])
        amino_max.append(data['patients_reference'].loc[1, as_name])
    
    # first four patients are printed, an empty columns follows and the next 4 patients are printed.
    for idx, (_, patient) in enumerate(data['data'].iterrows()):  # to get the line_number instead of the pandas index use enumerate here
        if idx % 8 == 0:
            idx_row = (idx//8)*gap_rows
            # write min max and amino names
            ws_patients.write(idx_row, 0, 'Normbereich', fmt.heading_left)
            ws_patients.write(idx_row+1, 0, 'min', fmt.heading_center)
            ws_patients.write(idx_row+1, 1, 'max', fmt.heading_center)
            ws_patients.write_column(idx_row+2, 0, amino_min, fmt.valid)
            ws_patients.write_column(idx_row+2, 1, amino_max, fmt.valid)
            ws_patients.write_column(idx_row+2, 2, amino_names, fmt.heading_right)
            ws_patients.write_column(idx_row+2, 12, amino_names, fmt.heading_left)
            
            for inv in invalids:
                pos = aoi.index(inv)
                pos += pos//3
                ws_patients.conditional_format(idx_row+2+pos, 12, idx_row+2+pos, 12, {'type': 'no_errors', 'format': fmt.invalid})
                ws_patients.conditional_format(idx_row+2+pos, 2, idx_row+2+pos, 2, {'type': 'no_errors', 'format': fmt.invalid})

            # set distance between row groups
            dist_rows = [6, 11, 16, 21, 26, 31]
            for r in dist_rows:
                ws_patients.set_row(r + (idx//8)*gap_rows, 10)
            # add page break
            # ws_patients.set_row((idx//8)*gap_rows - 1, 20)
            
        if idx % 4 == 0 and idx != 0:
            second_part = 1
        # after 8 patients a new page shall start
        if idx % 8 == 0 and idx != 0:
            #idx_row += gap_rows
            idx_col = offset_col
            second_part = 0
        
        # write patient id
        patient_id = patient.iloc[1]
        value = int(patient_id) if patient_id and patient_id.isdecimal() else patient_id
        # split too long text into a second line
        if len(patient_id) > 9 and isinstance(value, str):
            idx_spaces = [i for i, ltr in enumerate(value) if ltr == ' ']
            split_idx = 9
            for idx_s in idx_spaces:
                if idx_s >= 6:
                    split_idx = idx_s
                    break
            split_idx = min(split_idx, 10) + 1
            second_value = value[split_idx:]
            value = value[:split_idx]  # + '-'
            _logger.info(f'Splitted heading into two parts: {value}, {second_value}') 
            help_write(cfg, workbook, ws_patients, idx_row+1, idx_col+second_part, second_value, 3)  
        help_write(cfg, workbook, ws_patients, idx_row, idx_col+second_part, value, 3)   
        
        empty_row = 0
        offset = 2
        # print only aminos of interest
        for idx_as, as_name in enumerate(aoi):  # range(2, 22):                             
            if idx_as in add_empty_row:
                empty_row += 1
            val = round(patient.loc[as_name], 1)
            fmt_num = fmt_map.loc[fmt_map.index[idx], as_name]
            aminos_row_idx = idx_row+idx_as+empty_row+offset
            help_write(cfg, workbook, ws_patients, aminos_row_idx, idx_col+second_part, val, fmt_num)
            
        # go to the next patient slot
        idx_col += 1


def help_write(cfg, wb, ws, idx_r, idx_c, val, fmt_nr):
    if fmt_nr == -1:
        fmt = wb.add_format(cfg['format_number_invalid'])
    elif fmt_nr == 1:
        fmt = wb.add_format(cfg['format_number_low'])
    elif fmt_nr == 2:
        fmt = wb.add_format(cfg['format_number_high'])
    elif fmt_nr == 3:
        fmt = wb.add_format(cfg['format_heading'])
    else:
        fmt = wb.add_format(cfg['format_number_valid'])
    ws.write(idx_r, idx_c, val, fmt)


def write_maxtrix(idx_row, idx_col, data, worksheet, format_header):
    worksheet.write_row(idx_row, idx_col, data.columns.values.tolist()[1:], format_header)
    idx_row += 1
    data = data.fillna('NaN')
    for index, row in data.iterrows():
        worksheet.write_row(idx_row, idx_col, row.to_numpy()[1:])
        idx_row += 1    
    
    return idx_row


def exceltopdf(doc):
    excel = client.DispatchEx("Excel.Application")
    excel.Visible = 0

    wb = excel.Workbooks.Open(doc)
    export_path = os.path.splitext(doc)[0] + '.pdf'
    wb.WorkSheets([2, 3]).Select()
    try:
        wb.ActiveSheet.ExportAsFixedFormat(0, export_path)
        os.startfile(export_path)
    except Exception as e:
        _logger.critical(f'Failed to convert: {e}')
    finally:
        wb.Close()
        excel.Quit()
    
