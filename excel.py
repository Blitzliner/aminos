import xlsxwriter

def export(cfg, filename, raw_data, data, controls, checked_controls):
    workbook = xlsxwriter.Workbook(filename)
    ws_raw_data = workbook.add_worksheet('Rohdaten')
    ws_controls = workbook.add_worksheet('Kontrollen')
    ws_patients = workbook.add_worksheet('Patienten')
    
    cell_format_bold = workbook.add_format({'bold': True}) #, 'italic': True})
    
    # write the controls
    idx_row = 0
    idx_col = 0
    ws_raw_data.write_row(idx_row, idx_col, raw_data.columns.values.tolist()[1:], cell_format_bold)
    idx_row += 1
    for index, row in raw_data.iterrows():
        ws_raw_data.write_row(idx_row, idx_col, row.to_numpy()[1:])
        idx_row += 1
        
    # write the controls
    idx_row = 0
    idx_col = 0
    ws_controls.write_row(idx_row, idx_col, controls.columns.values.tolist()[1:], cell_format_bold)
    idx_row += 1
    for index, row in controls.iterrows():
        ws_controls.write_row(idx_row, idx_col, row.to_numpy()[1:])
        idx_row += 1

    idx_row += 2
    for index, row in checked_controls.iterrows():
        ws_controls.write_row(idx_row, idx_col, row.to_numpy()[1:])
        idx_row += 1
      
    #write the patient data
    idx_row = 0
    idx_col = 0
    ws_patients.write_row(idx_row, idx_col, data.columns.values.tolist()[1:], cell_format_bold)
    idx_row += 1
    for index, row in data.iterrows():
        ws_patients.write_row(idx_row, idx_col, row.to_numpy()[1:])
        idx_row += 1    
    
    
    workbook.close()