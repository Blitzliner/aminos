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
    if os.path.isfile(config_file):
        with open(config_file) as json_data_file:
            data = json.load(json_data_file)
    else:
        _logger.warning("config file does not exist. A default config file has been created.")
        data = {}
        data['export_directory'] = './analysed/'
        data['file_extension_raw_data'] = '_Rohdaten.xlsx'
        data['file_extension_analysis'] = '_Analyse.xlsx'
        data['ignore_samples'] = ['SIGMA200', 'SIGMA500', 'Phe200', 'Phe1000']
        
        with open(config_file, 'w') as fp:
            json.dump(data, fp)
            
    return data

def read_reference_data():
    _logger.info("read reference data")
    
def read_raw_data(filepath):
    data = {}
    if os.path.isfile(filepath):
        data = pd.read_excel(filepath) 
    else:
        _logger.error(F"could not read raw data. File is missing: {filepath}")
        
    return data

def get_timestamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
def preparation(cfg, raw_data_file):
    export_dir = cfg['export_directory']
    export_dir = os.path.join(export_dir, get_timestamp())
    if not os.path.isdir(export_dir):
        os.makedirs(export_dir)
    excel_sheet = get_timestamp() + cfg['file_extension_analysis']
    
    raw_copy_filename = get_timestamp() + cfg['file_extension_raw_data']
    
    copyfile(raw_data_file, os.path.join(export_dir, raw_copy_filename))
    
    return export_dir, excel_sheet
    

def main():
    _logger.info("start AMINOS tool")
    
    cfg = read_config()
    
    raw_data_file = 'rohdaten_example.xlsx'
    
    
    export_dir, excel_sheet_name = preparation(cfg, raw_data_file)
    excel_path = os.path.join(export_dir, excel_sheet_name)
    
    raw_data = read_raw_data(raw_data_file)
    
    print(raw_data.head())
    
    _logger.info("export to excel sheet")
    excel.export(excel_path)
    
    
    
if __name__== "__main__":
  main()