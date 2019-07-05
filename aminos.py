import json
import logging
import excel
import os
import datetime
import pandas as pd
from shutil import copyfile

#setup logger for console and file
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(name)s (%(lineno)s) - %(levelname)s: %(message)s", datefmt='%Y.%m.%d %H:%M:%S', filename="logger.log")

_logger = logging.getLogger("main")
logFormatter = logging.Formatter("%(asctime)s - %(name)s (%(lineno)s) - %(levelname)s: %(message)s")
consoleHandler = logging.StreamHandler()
consoleHandler.setFormatter(logFormatter)
consoleHandler.setLevel("INFO")
_logger.addHandler(consoleHandler)

#logging.basicConfig(format='%(asctime)s: %(levelname)s: %(message)s', level=logging.INFO)
#_logger = logging.getLogger("main")

def read_config(config_file = 'config.json', create_new_file=False):
    _logger.info("read config file")
    if create_new_file == False and os.path.isfile(config_file):
        with open(config_file) as json_data_file:
            data = json.load(json_data_file)
    else:
        _logger.warning("config file does not exist. A default config file has been created.")
        data = {}
        data['file_to_analyze'] = 'rohdaten_example.xlsx'
        data['export_directory'] = './analysed/'
        data['file_extension_raw_data'] = '_Rohdaten.xlsx'
        data['file_extension_analysis'] = '_Analyse.xlsx'
        data['ignore_samples'] = ['SIGMA200', 'SIGMA500', 'Phe200', 'Phe1000']
        data['control_name_prefix'] = 'Ko'
        data['control_ring_samples'] = [61, 62, 31, 32]
        data['max_normal_aminos'] = 21
        data['control_reference_file_path'] = 'kontrollwerte.csv'
        data['patients_reference_file_path'] = 'patienten_kontrollwerte.csv'
        data['format_heading'] = {'bold': True} #, 'bg_color': '#f1f2f6'
        data['format_number_invalid'] = {'bg_color': '#d1d8e0'}
        data['format_number_valid'] = {'bg_color': '#2bcbba'} ##26de81
        data['format_number_high'] = {'bg_color': '#fc5c65'}
        data['format_number_low'] = {'bg_color': '#45aaf2'} 
        data['prefer_control'] = 0
        data['prefer_aminos'] = []
        #warning-color: #fd9644, background: #4b6584
        columns = {}
        columns['sample_name'] = 'Sample Name'
        data['columns'] = columns
        
        with open(config_file, 'w') as fp:
            json.dump(data, fp, indent=4, sort_keys=True)
            
    return data

def read_reference_data(filepath):
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
    
def check_controls(cfg, data):
    controls = data['controls']
    control_reference = data['control_reference']
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

def select_control(cfg, data):
    dat = {}
    controls = data['controls']
    checked_controls = data['checked_controls']
    best_control = [0, 0]
    second_best_control = [0, 0]
    #max_prios_score = 0
    #best_control = 0
    dat['data'] = {}
    for ring in cfg['control_ring_samples']:
        column_name = cfg['columns']['sample_name']
        # split up the controls
        mask = controls[column_name].str.contains(str(ring))
        if any(mask):
            ring_data = {} 
            ring_data['data'] = controls[mask]
            ring_data['checked'] = checked_controls[mask]
            
            counts = ring_data['checked'].apply(pd.value_counts).fillna(0)
            ring_data['score'] = counts[(counts.index == 'NORMAL')]
            dat['data'][str(ring)] = ring_data
            
            ring_data['prios'], ring_data['conflicts'] = switch_amino_columns(cfg, ring, ring_data['score'])
            ring_data['prios_score'] = ring_data['prios'].sum(axis = 1, skipna = True).item()
            _logger.debug(ring_data['prios_score'])
            
            if best_control[1] < ring_data['prios_score']:
                second_best_control = best_control.copy()
                best_control[1] = ring_data['prios_score'] 
                best_control[0] = ring
    
    dat['best_control_score'] = best_control[1]
    dat['best_control_name'] = best_control[0]
    dat['second_best_control_score'] = second_best_control[1]
    dat['second_best_control_name'] = second_best_control[0]
    
    if dat['best_control_score'] == dat['best_control_name']:
        _logger.warnung("both controls does have the same score, took first")
    _logger.info(F"1. control: {str(dat['best_control_name'])}, score: {str(dat['best_control_score'])}")
    _logger.info(F"2. control: {str(dat['second_best_control_name'])}, score: {str(dat['second_best_control_score'])}")
        
    return dat

def switch_amino_columns(cfg, control, score):
    ret = pd.DataFrame().reindex_like(score)
    conflicts = []
    for col in score.columns:
        if (score.columns.get_loc(col) <= cfg['max_normal_aminos']):
            aminos = score.columns.str.contains(col)
            idx_name = score[score.columns[aminos]].idxmax(axis=1)
            # check for equal amino pairs 
            for col in score[score.columns[aminos]].columns:
                if col == idx_name.item():
                    continue
                if score[col].item() == score[idx_name.item()].item() and score[col].item() != 0:
                    # take preferation from settings
                    if len(cfg['prefer_aminos']) > 0:
                        if (col in cfg['prefer_aminos']): # switch to preferation
                            _logger.info(F"Take prefered AS from setting {col}")
                            idx_name = col
                        elif (idx_name.item() in cfg['prefer_aminos']):
                            _logger.info(F"Take prefered AS from setting: {idx_name.item()}")
                            idx_name = idx_name.item()
                        else:
                            _logger.warning(F"Prefered AS for control {str(control)} could not be found in settings file. {col} vs {idx_name.item()}")
                    else:
                        conflicts.append((idx_name.item(), col))
                        _logger.warning(F"Conflict in control {str(control)} with {idx_name.item()} and {col} (score {score[col].item()}), took {idx_name.item()}")
            ret[idx_name] = score[idx_name]
    
    return ret, conflicts       

def filter_patients_data(cfg, data):
    # overwrite 
    if cfg['prefer_control'] != 0:
        best_control = str(cfg['prefer_control'])
    else:
        best_control = str(data['selected_control']['best_control_name'])
    dat = data['selected_control']['data'][best_control]['prios']
    
    idx_not_null = dat.isnull() == False
    idx_zero = dat == 0
    
    _logger.info("sort out invalid AS from patients data") 
    idx_valids = []
    for idx, val in idx_not_null.T.iterrows(): 
        if val.item() == True:
            idx_valids.append(idx_not_null.columns.get_loc(idx))
    
    idx_invalids = []
    for idx, val in idx_zero.T.iloc[2:].iterrows(): 
        if val.item() == True:
            idx_invalids.append(idx_zero.columns.get_loc(idx))
    
    patients = data['data'].copy()
    #patients.loc[:, [idx_valids]]
    patients = patients[patients.columns[idx_valids]]
    
    _logger.info("sorting patients data") 
    lis = list(patients.columns.values)
    sorted_cols = lis[0:2] # ignore first two columns
    aminos_sorted = sorted(patients.columns[2:])
    sorted_cols.extend(aminos_sorted)
    new_patients = patients.reindex(sorted_cols, axis=1)
    return (new_patients, idx_invalids)

def analyse(cfg):
    _logger.info("start AMINOS tool")
    
    export_dir, excel_sheet_name = preparation(cfg, cfg['file_to_analyze'])
    excel_path = os.path.join(export_dir, excel_sheet_name)
    
    data = {}
    data['raw_data'] = read_raw_data(cfg['file_to_analyze'])
    data['data'], data['controls'] = filter_raw_data(cfg, data['raw_data'])
    data['control_reference'] = read_reference_data(cfg['control_reference_file_path'])
    data['patients_reference'] = read_reference_data(cfg['patients_reference_file_path'])
    data['checked_controls'] = check_controls(cfg, data)
    data['selected_control'] = select_control(cfg, data)
    data['data_filtered'], data['idx_invalids'] = filter_patients_data(cfg, data)
    
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