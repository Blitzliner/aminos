import xlsxwriter
import pandas as pd
import logging
_logger = logging.getLogger("excel")

def export(cfg, filename, data):    
    workbook = xlsxwriter.Workbook(filename)
    
    _logger.info("write raw data to excel")
    write_raw_data(workbook, data, cfg)
        
    _logger.info("write controls data to excel")
    write_controls_data(workbook, data, cfg)
   
    _logger.info("write patients data")
    write_patients_data(workbook, data, cfg)
    
    _logger.info("write control data")
    write_control_data(workbook, data, cfg)
    
    workbook.close()

def write_raw_data(workbook, data, cfg):
    ws_raw_data = workbook.add_worksheet('Rohdaten')
    ws_raw_data.set_landscape()
    fmt_heading = workbook.add_format(cfg['format_heading'])
    write_maxtrix(0, 0, data['raw_data'], ws_raw_data, format_header=fmt_heading)

def write_controls_data(workbook, data, cfg):
    ws_controls = workbook.add_worksheet('Kontrollen')    
    ws_controls.set_landscape()
    fmt_heading = workbook.add_format(cfg['format_heading'])
     
    splitted_controls = data['selected_control']['data']
    first = F"1. Wahl: Kontrolle: {str(data['selected_control']['best_control_name'])} Score: {str(data['selected_control']['best_control_score'])}"
    second = F"2. Wahl: Kontrolle: {str(data['selected_control']['second_best_control_name'])} Score: {str(data['selected_control']['second_best_control_score'])}"
    ws_controls.write(0, 0, first, fmt_heading)
    ws_controls.write(1, 0, second, fmt_heading)
    
    last_idx = 1
    for key in splitted_controls:
        ws_controls.write(last_idx+2, 0, "Rohdaten für Kontrolle: " + str(key), fmt_heading)
        last_idx = write_maxtrix(last_idx+3, 0, splitted_controls[str(key)]['data'], ws_controls, format_header=fmt_heading)
        
        ws_controls.write(last_idx+1, 0, "Bereichsanalyse", fmt_heading)
        last_idx = write_maxtrix(last_idx+2, 0, splitted_controls[str(key)]['checked'], ws_controls, format_header=fmt_heading)
        
        ws_controls.write(last_idx+1, 0, "Wie viele sind im Normbereich?", fmt_heading)
        last_idx = write_maxtrix(last_idx+2, 0, splitted_controls[str(key)]['score'], ws_controls, format_header=fmt_heading)
        
        ws_controls.write(last_idx+1, 0, "Priorisierung der AS", fmt_heading)
        prios = splitted_controls[str(key)]['prios']
        greater_zero = prios > 0
        prios[greater_zero] = 'okay'
        prios.replace(0, 'unpassend', inplace=True)
        prios = prios.fillna('ignoriert')
        last_idx = write_maxtrix(last_idx+2, 0, prios, ws_controls, format_header=fmt_heading)

def write_patients_data(workbook, data, cfg):    
    ws_patients = workbook.add_worksheet('Patienten')
    ws_patients.set_landscape()
    ws_patients.set_column("A:C", 4.5)
    ws_patients.set_column("I:I", 4.5)
    
    header3 = '&L&A' + '&CMessergebnisse des Aminosäure-Screenings' + '&RLSeite &P von &N'
    footer3 = '&RDatum: &D, &T'
    ws_patients.set_header(header3)
    ws_patients.set_footer(footer3)

    fmt_heading = workbook.add_format(cfg['format_heading'])
    fmt_heading_right = workbook.add_format(cfg['format_heading'])
    fmt_heading_right.set_align("right")
    # write additional infos to the the patient sheet
    #ws_patients.write(0,0,"Messergebnisse des Aminosäure-Screenings")
    ws_patients.write(0,0,"Normbereich", fmt_heading)
    ws_patients.write(1,0,"min", fmt_heading_right)
    ws_patients.write(1,1,"max", fmt_heading_right)
	    
    #format the patient data
    patients = data['data_filtered']
    pat_ref = data['patients_reference']
    fmt = pd.DataFrame().reindex_like(patients)
    
    for col in pat_ref.columns.values:
        val_min = pat_ref.loc[0,col]
        val_max = pat_ref.loc[1,col]
        col_name = patients.columns[patients.columns.str.contains(pat = col)][0]
        fmt[col_name][patients[col_name] < val_min] = 1 # mark as too low
        fmt[col_name][patients[col_name] > val_max] = 2 # mark as too high
        
    fmt[patients.columns[data['idx_invalids']]] = -1 # mark as AS invalid
    fmt.fillna(0, inplace=True) # mark rest as zero = valid
     
    gap_rows = 30
    offset_col = 4
    idx_row = 0
    idx_col = offset_col
    second_part = 0
    fmt_bold = workbook.add_format(cfg['format_heading'])
    #page_break = []
    # write min max and amino names 
    amino_names = []
    amino_min = []
    amino_max = []
    add_empty_row = [5, 8, 11, 14, 17, 20]
    for idx in range(2, 22):                             
        if idx in add_empty_row:
            amino_names.append("")
            amino_min.append("")
            amino_max.append("")
        as_name = data['data_filtered'].columns[idx]
        amino_names.append(as_name)
        amino_min.append(data['patients_reference'].loc[1, as_name[:3]])
        amino_max.append(data['patients_reference'].loc[0, as_name[:3]])
        
    ws_patients.write_column(idx_row+2, idx_col-1, amino_names, fmt_heading_right)
    ws_patients.write_column(idx_row+2, idx_col+9, amino_names, fmt_bold)
    ws_patients.write_column(idx_row+2, idx_col-3, amino_min)
    ws_patients.write_column(idx_row+2, idx_col-4, amino_max)
    
    for (idx, row) in data['data_filtered'].iterrows():
        # write patient id
        help_write(cfg, workbook, ws_patients, idx_row, idx_col+second_part, row.iloc[1], 3)    
        # write aminos for one patient
        empty_row = 0
        for idx_as in range(2, 22):                             
            if idx_as in add_empty_row:
                empty_row += 1
			val = F"{row.iloc[idx_as]:.1f}"
            help_write(cfg, workbook, ws_patients, idx_row+idx_as+empty_row, idx_col+second_part, val, fmt.loc[idx][idx_as])
            
        # go to the next patient slot
        idx_col += 1
        # first four patients are printed, an empty columns follows
        if idx_col%(4+offset_col) == 0: 
            second_part = 1
        # after 8 patients a new page shall start
        elif idx_col%(8+offset_col) == 0:
            idx_col = offset_col # start idx
            #page_break.append(idx_row)
            idx_row += gap_rows
            second_part = 0
            # write min max and amino names 
            ws_patients.write(idx_row,0,"Normbereich", fmt_heading)
            ws_patients.write(idx_row+1,0,"min", fmt_heading_right)
            ws_patients.write(idx_row+1,1,"max", fmt_heading_right)
            ws_patients.write_column(idx_row+2, idx_col-1, amino_names, fmt_heading_right )
            ws_patients.write_column(idx_row+2, idx_col+9, amino_names, fmt_bold )
            ws_patients.write_column(idx_row+2, idx_col-3, amino_min)
            ws_patients.write_column(idx_row+2, idx_col-4, amino_max)
    
    #ws_patients.h_pagebreaks = [20, 40, 80] #page_break
    #ws_patients.set_row(7, 30) 
    #ws_patients.set_row(12, 30) 
    #ws_patients.set_row(17, 30) 
    #ws_patients.set_row(22, 30) 
    
    ws_patients.set_row(28, 45) 
    ws_patients.set_row(29, 45) 
    ws_patients.set_row(58, 45) 
    ws_patients.set_row(59, 45) 
    ws_patients.set_row(88, 45) 
    ws_patients.set_row(89, 45) 
    
def write_control_data(wb, data, cfg):    
    ws = wb.add_worksheet('Gewählte Kontrolle')
    ws.set_landscape()
    header3 = '&L&A' + '&CMessergebnisse des Aminosäure-Screenings' + '&RSeite &P von &N'
    footer3 = '&RDatum: &D, &T'
    ws.set_header(header3)
    ws.set_footer(footer3)
    ws.set_column("A:A", 13)
    ws.set_column("B:Z", 4.5)
    fmt_heading = wb.add_format(cfg['format_heading'])
    fmt_heading.set_rotation(90)
    write_maxtrix(0, 0, data['control_filtered'], ws, format_header=fmt_heading)
         
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
        fmt = {}
    ws.write(idx_r, idx_c, val, fmt)
    
def write_maxtrix(idx_row, idx_col, data, worksheet, format_header):
    worksheet.write_row(idx_row, idx_col, data.columns.values.tolist()[1:], format_header)
    idx_row += 1
    data = data.fillna('NaN')
    for index, row in data.iterrows():
        worksheet.write_row(idx_row, idx_col, row.to_numpy()[1:])
        idx_row += 1    
    
    return idx_row