import json
import logging
import excel
import os
import datetime
import pandas as pd
from shutil import copyfile

logging.basicConfig(format='%(asctime)s: %(levelname)s: %(message)s', level=logging.INFO)
_logger = logging.getLogger("main")

def read_config(config_file = 'config.json'):
    _logger.info("read config file")
    #if os.path.isfile(config_file):
    #    with open(config_file) as json_data_file:
    #'        data = json.load(json_data_file)
    #'else:
    _logger.warning("config file does not exist. A default config file has been created.")
    data = {}
    data['export_directory'] = './analysed/'
    data['file_extension_raw_data'] = '_Rohdaten.xlsx'
    data['file_extension_analysis'] = '_Analyse.xlsx'
    data['ignore_samples'] = ['SIGMA200', 'SIGMA500', 'Phe200', 'Phe1000']
    data['control_name_prefix'] = 'Ko'
    data['control_ring_samples'] = [61, 62, 31, 32]
    data['control_reference_file_path'] = 'kontrollwerte.csv'
    columns = {}
    columns['sample_name'] = 'Sample Name'
    data['columns'] = columns
    
    with open(config_file, 'w') as fp:
        json.dump(data, fp, indent=4, sort_keys=True)
            
    return data

def read_control_reference_data(filepath):
    _logger.info("read control reference data")
    data = {}
    if os.path.isfile(filepath):
        data = pd.read_csv(filepath) 
    else:
        _logger.error(F"could not read control reference data. File is missing: {filepath}")    
    return data
    
def read_raw_data(filepath):
    _logger.info("read raw data")
    data = {}
    if os.path.isfile(filepath):
        data = pd.read_excel(filepath) 
    else:
        _logger.error(F"could not read raw data. File is missing: {filepath}")
    return data

def get_timestamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
def preparation(cfg, raw_data_file):
    _logger.info("prepare output directory and copy raw data")
    export_dir = cfg['export_directory']
    export_dir = os.path.join(export_dir, get_timestamp())
    if not os.path.isdir(export_dir):
        os.makedirs(export_dir)
    excel_sheet = get_timestamp() + cfg['file_extension_analysis']
    
    raw_copy_filename = get_timestamp() + cfg['file_extension_raw_data']
    
    copyfile(raw_data_file, os.path.join(export_dir, raw_copy_filename))
    
    return export_dir, excel_sheet
    
def filter_raw_data(cfg, data):
    _logger.info("ignore samples: " + ', '.join(cfg['ignore_samples']))
    column_name = cfg['columns']['sample_name']
    ignore = cfg['ignore_samples']
    control_name = cfg['control_name_prefix']
    data = data[data[column_name].isin(ignore) == False] # remove all unused samples from matrix
    controls_idx = data[column_name].str.contains(control_name, na=False) # get idx with controls
    
    _logger.info("split into controls and patient data")
    controls = data[controls_idx]                # get controls from matrix
    data = data[controls_idx == False]           # remove controls from data matrix
    
    _logger.info("sort data")
    controls.sort_values(column_name, axis=0, ascending=True, inplace=True) # sort ascending
    data.sort_values(column_name, axis=0, ascending=True, inplace=True) # sort ascending
    return data, controls
    
def check_controls(cfg, controls, control_reference):
    ret = pd.DataFrame().reindex_like(controls)
    ret[ret.isnull()] = 'NONE'
    
    str_ring_samples = ', '.join(str(s) for s in cfg['control_ring_samples'])
    _logger.info(F"check for following controls: {str_ring_samples}")
    
    for idx_row, row_data in controls.iterrows():
        # search in the controls for control ring samples
        control_name = row_data[cfg['columns']['sample_name']]
        matches = [num for num in cfg['control_ring_samples'] if str(num) in control_name]
        # check if a match exist
        if not len(matches):
            _logger.error('control can not be found in settings "control_ring_samples"')
        else:
            match = matches[0]
            _logger.info(F"processing control {match} for {control_name}")
            
            # get the limits from the reference frame
            matched_control = control_reference[control_reference['controls'] == match]
            control_min = matched_control[matched_control['limits'] == 'min']
            control_max = matched_control[matched_control['limits'] == 'max']
            
            for col in control_reference.columns.values:
                val_min = control_min[col].item() 
                val_max = control_max[col].item()
                
                col_names = controls.columns[controls.columns.str.contains(pat = col)]
                if (len(col_names)):
                
                    control_idx_bo = (controls[cfg['columns']['sample_name']] == control_name)
                    
                    for col_name in col_names:
                        ret[col_name][control_idx_bo & (controls[col_name] < val_min)] = 'TOO_LOW'
                        ret[col_name][control_idx_bo & (controls[col_name] > val_max)] = 'TOO_HIGH'
                        ret[col_name][control_idx_bo & (controls[col_name] <= val_max) & (controls[col_name] >= val_min)] = 'NORMAL'
                        
                        _logger.debug(F"{col} = {col_name}: {val_min} < {controls[col_name][control_idx_bo].to_string(index=False)} > {val_max}")
    return ret

def select_control(cfg, controls, checked_controls):
    dat = {}
    for ring in cfg['control_ring_samples']:
        column_name = cfg['columns']['sample_name']
        mask = controls[column_name].str.contains(str(ring))
        if any(mask):
            ring_data = {} 
            ring_data['data'] = controls[mask]
            ring_data['checked'] = checked_controls[mask]
            
            counts = ring_data['checked'].apply(pd.value_counts).fillna(0)
            ring_data['score'] = counts[(counts.index == 'NORMAL')]
            dat[str(ring)] = ring_data
            
    return dat

def main():
    _logger.info("start AMINOS tool")
    
    cfg = read_config()
    
    raw_data_file = 'rohdaten_example.xlsx'
    
    export_dir, excel_sheet_name = preparation(cfg, raw_data_file)
    excel_path = os.path.join(export_dir, excel_sheet_name)
    data = {}
    data['raw_data'] = read_raw_data(raw_data_file)
    data['data'], data['controls'] = filter_raw_data(cfg, data['raw_data'])
    data['control_reference'] = read_control_reference_data(cfg['control_reference_file_path'])
    data['checked_controls'] = check_controls(cfg, data['controls'], data['control_reference'])
    data['controls_counted'] = select_control(cfg, data['controls'], data['checked_controls'])
    
    print(data)
    #print(data['controls_counted'])
    #print(data['control_reference'])
    #print(data['checked_controls'])
    #print(data.head())#.head())
    #print(data['controls'])#.head())
    
    _logger.info("export to excel sheet")
    excel.export(cfg, excel_path, data['raw_data'], data['data'], data['controls'], data['checked_controls'])
    
    
if __name__== "__main__":
  main()