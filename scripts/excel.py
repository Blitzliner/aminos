import xlsxwriter
import pandas as pd
import logging
import os
import datetime
from win32com import client
import win32api

_logger = logging.getLogger("excel")

def export(cfg, filename, data):
    workbook = xlsxwriter.Workbook(filename)
    _logger.info("write raw data to excel")
    write_raw_data(workbook, data, cfg)
        
    _logger.info("write controls data to excel")
    write_controls_data(workbook, data, cfg)
   
    _logger.info("write patients data")
    write_patients_data(workbook, data, cfg)
   
    workbook.close()
    
    exceltopdf(filename)
    command = f"start EXCEL.EXE {filename}"
    os.system(command)

def write_raw_data(workbook, data, cfg):
    ws_raw_data = workbook.add_worksheet('Rohdaten')
    ws_raw_data.set_landscape()
    fmt_heading = workbook.add_format(cfg['format_heading'])
    write_maxtrix(0, 0, data['raw_data'], ws_raw_data, format_header=fmt_heading)

def write_controls_data(workbook, data, cfg):
    ws_controls = workbook.add_worksheet('Kontrollen')    
    ws_controls.set_landscape()
    ws_controls.set_header('&L&A' + '&CMessergebnisse des Aminosäure-Screenings' + '&RSeite &P von &N')
    ws_controls.set_footer('&RDatum: &D, &T')
    fmt_heading = workbook.add_format(cfg['format_heading'])
    fmt_heading_orient = workbook.add_format(cfg['format_heading'])
    fmt_heading_orient.set_rotation(90)
    fmt_heading_orient.set_align('center')
    fmt_center = workbook.add_format({'align': 'center'})
    fmt_even_row = workbook.add_format({'bg_color': '#eeeeee'})
    fmt_invalid = workbook.add_format(cfg['format_number_invalid'])
    
    ws_controls.write(0, 0, F"1. Wahl: Kontrolle {data['selected_control']}", fmt_heading)
    
    row_idx = 2
    ws_controls.write(row_idx, 0, 'Ranking', fmt_heading_orient)
    ws_controls.write(row_idx, 1, 'Kontrolle', fmt_heading_orient)
    ws_controls.write(row_idx, 2, 'Gültige AS', fmt_heading_orient)
    ws_controls.set_column('D:D', None, None, {'hidden': True})
    ws_controls.write(row_idx, 3, 'Score', fmt_heading_orient)
    ws_controls.write_row(row_idx, 5, data['controls'].columns.values.tolist()[2:], fmt_heading_orient)
    ws_controls.set_column('A:Y', 4.2)
    ws_controls.set_column('C:C', 4.8)
    ws_controls.set_column('E:E', 1.5)
    
    row_idx = 3
    first_invalid = None
    for rank, dat in enumerate(data['checked_controls']):
        ws_controls.write(row_idx, 0, rank+1, fmt_center)
        ws_controls.write(row_idx, 1, dat['name'], fmt_heading)
        ws_controls.write(row_idx, 2, f"{dat['coarse_score']}/20")
        ws_controls.write(row_idx, 3, f"{round(dat['fine_score']*100,1)}%", fmt_center)
        if rank%2 == 0:
            ws_controls.write_row(row_idx, 5, dat['raw_data'].to_numpy()[2:])
        else:
            ws_controls.write_row(row_idx, 5, dat['raw_data'].to_numpy()[2:], fmt_even_row)
        
        res = dat['result']
        # mark invalids 
        invalids = res[res=='TOO_LOW'].index.to_list() + res[res=='TOO_HIGH'].index.to_list()
        invalids = set(invalids)
        for inv in invalids:
            pos = res.index.get_loc(inv)
            ws_controls.conditional_format(row_idx, pos+5, row_idx, pos+5, {'type': 'no_errors', 'format': fmt_invalid}) 
            # mark heading as invalid
            if rank == 0:
                ws_controls.conditional_format(2, pos+5, 2, pos+5, {'type': 'no_errors', 'format': fmt_invalid}) 
        
        row_idx += 1
    
        
def write_patients_data(workbook, data, cfg):    
    ws_patients = workbook.add_worksheet('Patienten')
    ws_patients.set_landscape()
    ws_patients.set_header('&L&A' + '&CMessergebnisse des Aminosäure-Screenings' + '&RSeite &P von &N')
    ws_patients.set_footer('&RDatum: &D, &T')
    ws_patients.set_column("A:C", 5.2)
    ws_patients.set_column("I:I", 3.0)
    
    fmt_heading = workbook.add_format(cfg['format_heading'])
    fmt_heading.set_align('center')
    fmt_heading_right = workbook.add_format(cfg['format_heading'])
    fmt_heading_right.set_align("right")
    fmt_heading_left = workbook.add_format(cfg['format_heading'])
    fmt_heading_left.set_align("left")
    fmt_normal = workbook.add_format(cfg['format_number_valid'])
    fmt_invalid = workbook.add_format(cfg['format_number_invalid'])

    #format the patient data
    patients = data['data']
    pat_ref = data['patients_reference']
    fmt = pd.DataFrame().reindex_like(patients)
    fmt.iloc[:] = 0  # init with zeros
    fmt = fmt.astype('int')  # convert to integer
    
    # iterate by aminos
    for col in pat_ref.columns.values:
        val_min, val_max = pat_ref.loc[:, col]
        fmt[col][patients[col] < val_min] = 1 # mark as too low
        fmt[col][patients[col] > val_max] = 2 # mark as too high
    # mark aminos as invalid 
    res = data['checked_controls'][0]['result']
    invalids = res[res=='TOO_LOW'].index.to_list() + res[res=='TOO_HIGH'].index.to_list()
    invalids = list(set(invalids))
    fmt[invalids] = -1

    # 
    gap_rows = 30
    offset_col = 4
    idx_row = 0
    idx_col = offset_col
    second_part = 0
    fmt_bold = workbook.add_format(cfg['format_heading'])
    
    # reformat array with data. Insert empty column
    amino_names = []
    amino_min = []
    amino_max = []
    add_empty_row = [5, 8, 11, 14, 17, 20]
    for idx in range(2, 22):                             
        if idx in add_empty_row:
            amino_names.append('')
            amino_min.append('')
            amino_max.append('')
        as_name = data['data'].columns[idx]
        amino_names.append(as_name)
        amino_min.append(data['patients_reference'].loc[1, as_name])
        amino_max.append(data['patients_reference'].loc[0, as_name])
        
    # first four patients are printed, an empty columns follows and the next 4 patients are printed.
    for (idx, patient) in data['data'].iterrows():
        if idx%8 == 0:
            idx_row = idx//8*gap_rows
            # write min max and amino names
            ws_patients.write(idx_row, 0, "Normbereich", fmt_heading_left)
            ws_patients.write(idx_row+1, 0, "min", fmt_heading)
            ws_patients.write(idx_row+1, 1, "max", fmt_heading)
            ws_patients.write_column(idx_row+2, 0, amino_min, fmt_normal)
            ws_patients.write_column(idx_row+2, 1, amino_max, fmt_normal)
            ws_patients.write_column(idx_row+2, 3, amino_names, fmt_heading_right)
            ws_patients.write_column(idx_row+2, 13, amino_names, fmt_heading_left)
            for inv in invalids:
                pos = res.index.get_loc(inv)
                pos += pos//3
                ws_patients.conditional_format(idx_row+2+pos, 13, idx_row+2+pos, 13, {'type': 'no_errors', 'format': fmt_invalid})
                ws_patients.conditional_format(idx_row+2+pos, 3, idx_row+2+pos, 3, {'type': 'no_errors', 'format': fmt_invalid})
            
            # add page break
            offset = gap_rows-1
            ws_patients.set_row(offset + (idx//8)*gap_rows - 1, 45) 
            ws_patients.set_row(offset + (idx//8)*gap_rows, 45)
            
        if idx%4 == 0 and idx != 0:
            second_part = 1
        # after 8 patients a new page shall start
        if idx%8 == 0 and idx != 0:
            #idx_row += gap_rows
            idx_col = offset_col
            second_part = 0
        
        # write patient id
        help_write(cfg, workbook, ws_patients, idx_row, idx_col+second_part, int(patient.iloc[1]), 3)   

        # write aminos for one patient
        empty_row = 0
        for idx_as in range(2, 22):                             
            if idx_as in add_empty_row:
                empty_row += 1
            val = patient.iloc[idx_as]
            help_write(cfg, workbook, ws_patients, idx_row+idx_as+empty_row, idx_col+second_part, val, fmt.loc[idx][idx_as])
        
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
        #wb.Worksheets[1:].SaveAs(export_path, FileFormat=57)
        wb.ActiveSheet.ExportAsFixedFormat(0, export_path)
    except Exception as e:
        print(f"Failed to convert: {e}")
    finally:
        wb.Close()
        excel.Quit()