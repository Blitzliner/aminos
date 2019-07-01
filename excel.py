import xlsxwriter
import pandas as pd

def export(cfg, filename, data):
    raw_data = data['raw_data']
    cleaned_data = data['data']
    controls = data['controls']
    checked_controls = data['checked_controls']
    workbook = xlsxwriter.Workbook(filename)
    
    ws_raw_data = workbook.add_worksheet('Rohdaten')
    ws_controls = workbook.add_worksheet('Kontrollen')
    ws_patients = workbook.add_worksheet('Patienten')
    
    cell_format_bold = workbook.add_format({'bold': True}) #, 'italic': True})
    
    # write raw data
    write_maxtrix(0, 0, raw_data, ws_raw_data, format_header=cell_format_bold)
    
    # write controls
    #last_idx = write_maxtrix(0, 0, controls, ws_controls, format_header=cell_format_bold)
    #last_idx = 0
    splitted_controls = data['selected_control']['data']
    first = F"1. Wahl: Kontrolle: {str(data['selected_control']['best_control_name'])} Score: {str(data['selected_control']['best_control_score'])}"
    second = F"1. Wahl: Kontrolle: {str(data['selected_control']['second_best_control_name'])} Score: {str(data['selected_control']['second_best_control_score'])}"
    ws_controls.write(0, 0, first, cell_format_bold)
    ws_controls.write(1, 0, second, cell_format_bold)
    #ws_controls.write(2, 0, "Score: " + str(data['selected_control']['best_control_score']))
    #ws_controls.write(4, 0, "2. Wahl: ", cell_format_bold)
    #ws_controls.write(5, 0, "Kontrolle: " + str(data['selected_control']['second_best_control_name']))
    #ws_controls.write(6, 0, "Score: " + str(data['selected_control']['second_best_control_score']))
    last_idx = 1
    for key in splitted_controls:
        ws_controls.write(last_idx+2, 0, "Rohdaten für Kontrolle: " + str(key), cell_format_bold)
        last_idx = write_maxtrix(last_idx+3, 0, splitted_controls[str(key)]['data'], ws_controls, format_header=cell_format_bold)
        
        ws_controls.write(last_idx+1, 0, "Bereichsanalyse", cell_format_bold)
        last_idx = write_maxtrix(last_idx+2, 0, splitted_controls[str(key)]['checked'], ws_controls, format_header=cell_format_bold)
        
        ws_controls.write(last_idx+1, 0, "Wie viele sind im Normbereich?", cell_format_bold)
        last_idx = write_maxtrix(last_idx+2, 0, splitted_controls[str(key)]['score'], ws_controls, format_header=cell_format_bold)
        
        ws_controls.write(last_idx+1, 0, "Priorisierung der AS", cell_format_bold)
        prios = splitted_controls[str(key)]['prios']
        greater_zero = prios > 0
        equal_zero = prios == 0
        prios[greater_zero] = 'okay'
        
        prios.replace(0, 'unpassend', inplace=True)
        #prios[equal_zero] = 'unpassend'
        prios = prios.fillna('ignoriert')
        last_idx = write_maxtrix(last_idx+2, 0, prios, ws_controls, format_header=cell_format_bold)
        #print( data['selected_control'][key])
        
    #########################################################################
    # Patients sheet
    #########################################################################
    # write additional infos to the the patient sheet
    ws_patients.write(0,0,"Messergebnisse des Aminosäure-Screenings")
    ws_patients.write(1,0,"Normbereich", cell_format_bold)
    ws_patients.write(2,0,"min", cell_format_bold)
    ws_patients.write(2,1,"max", cell_format_bold)
	    
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
     
    gap_rows = 26
    offset_col = 4
    idx_row = 2
    idx_col = offset_col
    second_part = 0
    fmt_bold = workbook.add_format({'bold': True})
    ws_patients.write_column(idx_row+2, idx_col-1, data['data_filtered'].columns.values.tolist()[2:], fmt_bold )
    ws_patients.write_column(idx_row+2, idx_col+9, data['data_filtered'].columns.values.tolist()[2:], fmt_bold )
    ws_patients.write_column(idx_row+2, idx_col-3, data['patients_reference'].loc[1,:])
    ws_patients.write_column(idx_row+2, idx_col-4, data['patients_reference'].loc[0,:])
    
    for (idx, row) in data['data_filtered'].iterrows():
        help_write(workbook, ws_patients, idx_row, idx_col+second_part, row.iloc[1], 3)    # patient id
        for idx_as in range(2, 22):                             # write aminos
            help_write(workbook, ws_patients, idx_row+idx_as, idx_col+second_part, row.iloc[idx_as], fmt.loc[idx][idx_as])
        
        idx_col += 1
        if idx_col%(4+offset_col) == 0: # first four patients are printed, an empty columns follows
            second_part = 1
        elif idx_col%(8+offset_col) == 0:
            idx_col = offset_col # start idx
            idx_row += gap_rows
            second_part = 0
            ws_patients.write_column(idx_row+2, idx_col-1, data['data_filtered'].columns.values.tolist()[2:], fmt_bold )
            ws_patients.write_column(idx_row+2, idx_col+9, data['data_filtered'].columns.values.tolist()[2:], fmt_bold )
            ws_patients.write_column(idx_row+2, idx_col-3, data['patients_reference'].loc[1,:])
            ws_patients.write_column(idx_row+2, idx_col-4, data['patients_reference'].loc[0,:])
    
    
    workbook.close()

def help_write(wb, ws, idx_r, idx_c, val, fmt_nr):
    if fmt_nr == -1:
        fmt = wb.add_format({'bg_color': 'gray'})
    elif fmt_nr == 1:
        fmt = wb.add_format({'bg_color': 'red'})
    elif fmt_nr == 2:
        fmt = wb.add_format({'bg_color': 'green'})
    elif fmt_nr == 3:
        fmt = wb.add_format({'bold': True})
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