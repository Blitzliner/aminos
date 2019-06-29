import json
import logging
import excel
import os
import datetime
logging.basicConfig(format='%(asctime)s: %(levelname)s: %(message)s', level=logging.INFO)
_logger = logging.getLogger("main")


def load_config(config_file = 'config.json'):
    if os.path.isfile(config_file):
        with open(config_file) as json_data_file:
            data = json.load(json_data_file)
    else:
        _logger.warning("config file does not exist. A default config file has been created.")
        data = {}
        data['export_directory'] = './analysed/'
        
        with open(config_file, 'w') as fp:
            json.dump(data, fp)
            
    return data

def get_timestamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
def get_export_path(export_dir):
    export_dir = os.path.join(export_dir, get_timestamp())
    if not os.path.isdir(export_dir):
        os.makedirs(export_dir)
    return export_dir
    
    
def main():
    _logger.info("start AMINOS tool")
    
    cfg = load_config()
    
    export_dir = get_export_path(cfg['export_directory'])
    excel_sheet = get_timestamp() + '.xlsx'
    excel_path = os.path.join(export_dir, excel_sheet)
    
    excel.export(excel_path)
    
if __name__== "__main__":
  main()