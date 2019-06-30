import xlsxwriter

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
        
    #write the patient data
    last_idx = write_maxtrix(0, 0, cleaned_data, ws_patients, format_header=cell_format_bold)
    last_idx = write_maxtrix(last_idx + 3, 0, data['data_filtered'], ws_patients, format_header=cell_format_bold)
    
    workbook.close()
    
def write_maxtrix(idx_row, idx_col, data, worksheet, format_header):
    worksheet.write_row(idx_row, idx_col, data.columns.values.tolist()[1:], format_header)
    idx_row += 1
    data = data.fillna('NaN')
    for index, row in data.iterrows():
        worksheet.write_row(idx_row, idx_col, row.to_numpy()[1:])
        idx_row += 1    
    
    return idx_row