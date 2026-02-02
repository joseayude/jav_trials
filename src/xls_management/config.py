import yaml
from xls_management import ROOTPATH, HOMEPATH

class ATEConfig():
    config_file = ROOTPATH / "config.yml"
    def __init__(self):
        self.config:dict|None = None
        if(ATEConfig.config_file.exists()):
            self.load_config_file()
        else:
            self.set_default_config_file()

    def load_config_file(self) -> None:
        with open(ATEConfig.config_file, 'r') as file:
            self.config = yaml.safe_load(file)

    def set_default_config_file(self) -> None:
        self.config = {}
        self.config['BsM_workbook'] = (
            HOMEPATH / 'vw/data/ATE-Status_Berichtsversion.xlsm'
        ).as_uri()
        yaml_str = yaml.dump(self.config)
        with open(ATEConfig.config_file, 'w') as file:
            file.writelines(yaml_str)
