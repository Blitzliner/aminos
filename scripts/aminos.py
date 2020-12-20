import json
import logging
import excel
import os
import datetime
import pandas as pd
import re
from shutil import copyfile

#setup logger for console and file
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(name)s (%(lineno)s) - %(levelname)s: %(message)s", datefmt='%Y.%m.%d %H:%M:%S', filename="logger.log")

_logger = logging.getLogger("main")
logFormatter = logging.Formatter("%(asctime)s - %(name)s (%(lineno)s) - %(levelname)s: %(message)s")
consoleHandler = logging.StreamHandler()
consoleHandler.setFormatter(logFormatter)
consoleHandler.setLevel("INFO")
_logger.addHandler(consoleHandler)


def read_config(config_file = 'config.json', create_new_file=False):
    _logger.info("read config file")
    if create_new_file == False and os.path.isfile(config_file):
        with open(config_file) as json_data_file:
            data = json.load(json_data_file)
    else:
        _logger.warning("config file does not exist. A default config file has been created.")
        data = {}
        data['file_to_analyze'] = '../rohdaten_example.xlsx'
        data['export_directory'] = '../analysed/'
        data['file_extension_raw_data'] = '_Rohdaten.xlsx'
        data['file_extension_analysis'] = '_Analyse.xlsx'
        data['ignore_calibration'] = "[KkCc]al\s?\d"
        data['control_name'] = '(([CckK][oO])|([qQ][cC]))\\s?[I\\d]'
        data['control_ring_samples'] = [61, 62, 31, 32]
        data['max_normal_aminos'] = 21
        data['control_reference_file_path'] = './reference/kontrollwerte.csv'
        data['patients_reference_file_path'] = './reference/patienten_kontrollwerte.csv'
        data['format_heading'] = {'bold': True} #, 'bg_color': '#f1f2f6'
        data['format_number_invalid'] = {'bg_color': '#d1d8e0', 'font_color': '#636e72'}
        data['format_number_valid'] = {'bg_color': '#2bcbba'} ##26de81
        data['format_number_high'] = {'bg_color': '#fc5c65'}
        data['format_number_low'] = {'bg_color': '#45aaf2'} 
        data['prefer_control'] = 0
        #warning-color: #fd9644, background: #4b6584
        columns = {}
        columns['sample_name'] = 'Sample Name'
        data['columns'] = columns
        
        with open(config_file, 'w') as fp:
            json.dump(data, fp, indent=4, sort_keys=True)
            
    return data

def read_reference_data(filepath):
    filepath = os.path.abspath(filepath)
    _logger.info(F"read {filepath} reference data")
    data = {}
    if os.path.isfile(filepath):
        data = pd.read_csv(filepath) 
    else:
        _logger.error(F"could not read reference data. File is missing: {filepath}")    
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
    timestamp = get_timestamp()
    export_dir = cfg['export_directory']
    export_dir = os.path.join(export_dir, timestamp)
    if not os.path.isdir(export_dir):
        os.makedirs(export_dir)
    excel_sheet = timestamp + cfg['file_extension_analysis']
    
    raw_copy_filename = timestamp + cfg['file_extension_raw_data']
    
    copyfile(raw_data_file, os.path.join(export_dir, raw_copy_filename))
    
    return export_dir, excel_sheet

def filter_raw_data(cfg, data):
    _logger.info(f"remove calibration from data: {cfg['ignore_calibration']}")
    column_name = cfg['columns']['sample_name']
    ignore_pattern = cfg['ignore_calibration']
    control_name_pattern = cfg['control_name']
    data = data.drop(data[data[column_name].str.match(ignore_pattern) == True].index) # remove all unused calibration
    controls_idx = (data[column_name].str.match(control_name_pattern) == True)
    _logger.info("split into controls and patient data")
    controls = data[controls_idx]                # get controls from matrix
    data = data[controls_idx == False]           # remove controls from data matrix
    _logger.info(f"set invalid values 'No Peak' to None")
    controls = controls.replace("No Peak", None)
    data = data.replace("No Peak", None)
    _logger.info("sort data")
    controls[column_name] = controls[column_name].astype(str)
    controls.sort_values(column_name, axis=0, ascending=True, inplace=True) # sort ascending
    data[column_name] = data[column_name].astype(str)
    data.sort_values(column_name, axis=0, ascending=True, inplace=True) # sort ascending
    return data, controls
    
def check_controls(cfg, data):
    controls = data['controls']
    control_reference = data['control_reference']
    results = []
    
    for idx_row, row_data in controls.iterrows():
        # search in the controls for control ring samples
        control_name = row_data[cfg['columns']['sample_name']]
        if 'I' in control_name:
            raw_num = control_name.count('I')  # convert III integer
        else:
            raw_num = int(re.search(r'(\d)', control_name)[0])
        _logger.info(F"Get reference control dataset {raw_num} for control {control_name}")
        
        # get the limits from the reference frame
        matched_control = control_reference[control_reference['controls'] == raw_num]
        
        meas_con = row_data[2:]
        ref_min = matched_control.iloc[0, 2:]
        ref_mean = matched_control.iloc[2, 2:]
        ref_max = matched_control.iloc[2, 2:]
        
        # get the shape and initialize with NORMAL
        res = meas_con.copy()
        res.iloc[:] = 'NORMAL'
        # mark as too high or too low
        res[meas_con < ref_min] = 'TOO_LOW'
        res[meas_con > ref_max] = 'TOO_HIGH'
        
        # get coarse score
        too_high_count = res[res=='TOO_HIGH'].count()
        too_low_count = res[res=='TOO_LOW'].count()
        coarse_score = 20 - too_high_count - too_low_count
        
        # get fine score 
        error = (meas_con - ref_mean).abs()
        delta = (ref_mean - ref_min).abs()
        fine_score = round((1 - error/delta).mean(), 3)
        
        results.append({'name': control_name, 'result': res, 'coarse_score': coarse_score, 'fine_score': fine_score, 'raw_data': row_data})
        _logger.info(f'{control_name} score: {coarse_score}/{fine_score}')
        
    return results

def select_control(data):
    data['checked_controls'] = sorted(data['checked_controls'], key=lambda i: (i['coarse_score'], i['fine_score']), reverse=True)
    best = data['checked_controls'][0]
    _logger.info(F'best control: {best["name"]} with score {best["coarse_score"]}/{best["fine_score"]}')  
    return best['name']

def analyse(cfg):
    _logger.info("start AMINOS tool")
    data = {}
    
    export_dir, excel_sheet_name = preparation(cfg, cfg['file_to_analyze'])
    export_dir = os.path.abspath(export_dir)
    excel_path = os.path.abspath(os.path.join(export_dir, excel_sheet_name))
    _logger.info(export_dir)
    _logger.info(excel_path)
    
    data['export_dir'] = export_dir
    data['export_excel_path'] = excel_path
    data['raw_data'] = read_raw_data(cfg['file_to_analyze'])
    data['data'], data['controls'] = filter_raw_data(cfg, data['raw_data'])
    data['control_reference'] = read_reference_data(cfg['control_reference_file_path'])
    data['patients_reference'] = read_reference_data(cfg['patients_reference_file_path'])
    data['checked_controls'] = check_controls(cfg, data)
    data['selected_control'] = select_control(data)
    
    # temporaly write into file
    # with open('data.pickle', 'wb') as handle:
    #    pickle.dump(data, handle)
    #_logger.debug(data)
    
    excel.export(cfg, excel_path, data)
    
    _logger.info("finished analyses")
    
    return data
    
if __name__== "__main__":
    cfg = read_config()
    analyse(cfg)