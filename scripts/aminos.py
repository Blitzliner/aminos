import json
import logging
import excel
import os
import datetime
import pandas as pd
import re
from shutil import copyfile
from collections import Counter

#setup logger for console and file
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(name)s (%(lineno)s) - %(levelname)s: %(message)s", datefmt='%Y.%m.%d %H:%M:%S', filename="logger.log")

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
        data['control_reference_file_path'] = './reference/kontrollwerte.csv'
        data['patients_reference_file_path'] = './reference/patienten_kontrollwerte.csv'
        data['format_heading'] = {'bold': True}
        data['format_number_invalid'] = {'bg_color': '#cccccc', 'font_color': '#666666', 'align': 'center'}
        data['format_number_valid'] = {'font_color': '#000000', 'align': 'center'}
        data['format_number_high'] = {'bg_color': '#FAAAAA', 'font_color': '#A03232', 'align': 'center'}
        data['format_number_low'] = {'bg_color': '#c8c8fa', 'font_color': '#323250', 'align': 'center'} 
        data['prefer_control'] = ''
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
    remove_cols_idx = data[data[column_name].str.match(ignore_pattern) == True].index
    col_names = data.loc[remove_cols_idx, column_name]
    _logger.info(f'Remove columns: {", ".join(col_names)}')
    data = data.drop(remove_cols_idx) # remove all unused calibration
    controls_idx = (data[column_name].str.match(control_name_pattern) == True)
    _logger.info("split into controls and patient data")
    controls = data[controls_idx]  # get controls from matrix
    data = data[controls_idx == False]  # remove controls from data matrix
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
            control_num = control_name.count('I')  # convert III integer
        else:
            control_num = int(re.search(r'(\d)', control_name)[0])
        _logger.info(F"Get reference control dataset {control_num} for control {control_name}")
        
        # get the limits from the reference frame
        matched_control = control_reference[control_reference['controls'] == control_num]
        
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
    
    # change name of control if control name already exist
    all_control_names = [control['name'] for control in results]
    # first count number of occurences
    controls_count = Counter(all_control_names)
    # after the have the number of each control occurence we manipulate the name
    for control in results:
        name = control['name']
        count = controls_count[name]  # all_control_names.count(name)  # 
        control['name'] += f'_{count}'
        controls_count[name] -= 1
    return results

def select_control(cfg, data):
    prefered_control = cfg['prefer_control']
    data['checked_controls'] = sorted(data['checked_controls'], key=lambda i: (i['coarse_score'], i['fine_score']), reverse=True)
    best = data['checked_controls'][0]
    _logger.info(F'best control: {best["name"]} with score {best["coarse_score"]}/{best["fine_score"]}')
    
    if len(cfg['prefer_control']) > 0:
        if any(cfg['prefer_control'] == control['name'] for control in data['checked_controls']):
            _logger.info(f"Your prefered control: {cfg['prefer_control']}")
        else:
            prefered_control = best['name']
            _logger.warning(f"Your prefered control seems not to be valid: {cfg['prefer_control']}")
            _logger.info(f"I use the best control for you: {prefered_control}")
    else:
        prefered_control = best['name']
        _logger.info(f"No prefered control selected. I use the best for you: {prefered_control}")
    
    prefered_control_dat = [dat for dat in data['checked_controls'] if prefered_control == dat['name']][0]
    
    return prefered_control_dat

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
    data['selected_control'] = select_control(cfg, data)
    
    excel.export(cfg, excel_path, data)
    
    _logger.info("finished analyses")
    
    return data
    
if __name__== "__main__":
    cfg = read_config()
    analyse(cfg)