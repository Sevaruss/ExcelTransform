import configparser


class ConfigParams:
    _file_name = 'ExcelTransform.ini'

    def __init__(self):
        config = configparser.ConfigParser()
        config.read(self._file_name, encoding='utf-8')
        # print("Sections:", config.sections())

        # region Initialize all members
        self.UvedFileName = config.get('File', 'UvedFileName')
        self.SvodFileName = config.get('File', 'SvodFileName')
        #################################################################
        self.DUvedWSName = config.get('Uvedomlenie', 'DUvedWSName')
        self.DUvedPrefix = config.get('Uvedomlenie', 'DUvedPrefix')
        self.DUvedRowStart = int(config.get('Uvedomlenie', 'DUvedRowStart'))
        self.DUvedRowEnd = int(config.get('Uvedomlenie', 'DUvedRowEnd'))
        self.DUvedCollStart = int(config.get('Uvedomlenie', 'DUvedCollStart'))
        self.DUvedCollEnd = int(config.get('Uvedomlenie', 'DUvedCollEnd'))

        self.NPUvedWSName = config.get('Uvedomlenie', 'NPUvedWSName')
        self.NPUvedPrefix = config.get('Uvedomlenie', 'NPUvedPrefix')
        self.NPUvedRowStart = int(config.get('Uvedomlenie', 'NPUvedRowStart'))
        self.NPUvedRowEnd = int(config.get('Uvedomlenie', 'NPUvedRowEnd'))
        self.NPUvedCollStart = int(config.get('Uvedomlenie', 'NPUvedCollStart'))
        self.NPUvedCollEnd = int(config.get('Uvedomlenie', 'NPUvedCollEnd'))

        self.PUvedWSName = config.get('Uvedomlenie', 'PUvedWSName')
        self.PUvedPrefix = config.get('Uvedomlenie', 'PUvedPrefix')
        self.PUvedRowStart = int(config.get('Uvedomlenie', 'PUvedRowStart'))
        self.PUvedRowEnd = int(config.get('Uvedomlenie', 'PUvedRowEnd'))
        self.PUvedCollStart = int(config.get('Uvedomlenie', 'PUvedCollStart'))
        self.PUvedCollEnd = int(config.get('Uvedomlenie', 'PUvedCollEnd'))

        self.T2UvedWSName = config.get('Uvedomlenie', 'T2UvedWSName')
        self.T2UvedPrefix = config.get('Uvedomlenie', 'T2UvedPrefix')
        self.T2UvedRowStart = int(config.get('Uvedomlenie', 'T2UvedRowStart'))
        self.T2UvedRowEnd = int(config.get('Uvedomlenie', 'T2UvedRowEnd'))
        self.T2UvedCollStart = int(config.get('Uvedomlenie', 'T2UvedCollStart'))
        self.T2UvedCollEnd = int(config.get('Uvedomlenie', 'T2UvedCollEnd'))

        #############################################################
        self.DSvodWSName = config.get('Svod', 'DSvodWSName')
        self.DSvodPrefix = config.get('Svod', 'DSvodPrefix')
        self.DSvodRowStart = int(config.get('Svod', 'DSvodRowStart'))
        self.DSvodRowEnd = int(config.get('Svod', 'DSvodRowEnd'))
        self.DSvodCollStart = int(config.get('Svod', 'DSvodCollStart'))
        self.DSvodCollEnd = int(config.get('Svod', 'DSvodCollEnd'))

        self.NPSvodWSName = config.get('Svod', 'NPSvodWSName')
        self.NPSvodPrefix = config.get('Svod', 'NPSvodPrefix')
        self.NPSvodRowStart = int(config.get('Svod', 'NPSvodRowStart'))
        self.NPSvodRowEnd = int(config.get('Svod', 'NPSvodRowEnd'))
        self.NPSvodCollStart = int(config.get('Svod', 'NPSvodCollStart'))
        self.NPSvodCollEnd = int(config.get('Svod', 'NPSvodCollEnd'))

        self.PSvodWSName = config.get('Svod', 'PSvodWSName')
        self.PSvodPrefix = config.get('Svod', 'PSvodPrefix')
        self.PSvodRowStart = int(config.get('Svod', 'PSvodRowStart'))
        self.PSvodRowEnd = int(config.get('Svod', 'PSvodRowEnd'))
        self.PSvodCollStart = int(config.get('Svod', 'PSvodCollStart'))
        self.PSvodCollEnd = int(config.get('Svod', 'PSvodCollEnd'))

        self.T1SvodWSName = config.get('Svod', 'T1SvodWSName')
        self.T1SvodPrefix = config.get('Svod', 'T1SvodPrefix')
        self.T1SvodRowStart = int(config.get('Svod', 'T1SvodRowStart'))
        self.T1SvodRowEnd = int(config.get('Svod', 'T1SvodRowEnd'))
        self.T1SvodCollStart = int(config.get('Svod', 'T1SvodCollStart'))
        self.T1SvodCollEnd = int(config.get('Svod', 'T1SvodCollEnd'))

        self.T2SvodWSName = config.get('Svod', 'T2SvodWSName')
        self.T2SvodPrefix = config.get('Svod', 'T2SvodPrefix')
        self.T2SvodRowStart = int(config.get('Svod', 'T2SvodRowStart'))
        self.T2SvodRowEnd = int(config.get('Svod', 'T2SvodRowEnd'))
        self.T2SvodCollStart = int(config.get('Svod', 'T2SvodCollStart'))
        self.T2SvodCollEnd = int(config.get('Svod', 'T2SvodCollEnd'))

        # endregion

    def create_ini_file(self):
        config = configparser.ConfigParser()

        # region Fill config members with default values
        config['File'] = {
            'UvedFileName': 'svodUvedomlenie.txt',
            'SvodFileName': 'svod2020.txt'
        }

        config['Uvedomlenie'] = {
            'DUvedWSName': 'Документ - Уведомление',
            'DUvedPrefix': 'D',
            'DUvedRowStart': '5',
            'DUvedRowEnd': '5',
            'DUvedCollStart': '2',
            'DUvedCollEnd': '14',

            'NPUvedWSName': 'СведенияНП - Уведомление',
            'NPUvedPrefix': 'NP',
            'NPUvedRowStart': '5',
            'NPUvedRowEnd': '5',
            'NPUvedCollStart': '2',
            'NPUvedCollEnd': '15',

            'PUvedWSName': 'Подписант - Уведомление',
            'PUvedPrefix': 'P',
            'PUvedRowStart': '5',
            'PUvedRowEnd': '5',
            'PUvedCollStart': '2',
            'PUvedCollEnd': '10',

            'T2UvedWSName': 'Уведомление',
            'T2UvedPrefix': 'T2',
            'T2UvedRowStart': '6',
            'T2UvedRowEnd': '-1',
            'T2UvedCollStart': '2',
            'T2UvedCollEnd': '15'
        }

        config['Svod'] = {
            'DSvodWSName': 'Документ - Страновой',
            'DSvodPrefix': 'D',
            'DSvodRowStart': '5',
            'DSvodRowEnd': '5',
            'DSvodCollStart': '2',
            'DSvodCollEnd': '19',

            'NPSvodWSName': 'СведенияНП - Страновой',
            'NPSvodPrefix': 'NP',
            'NPSvodRowStart': '5',
            'NPSvodRowEnd': '5',
            'NPSvodCollStart': '2',
            'NPSvodCollEnd': '34',

            'PSvodWSName': 'Подписант - Страновой',
            'PSvodPrefix': 'P',
            'PSvodRowStart': '5',
            'PSvodRowEnd': '5',
            'PSvodCollStart': '2',
            'PSvodCollEnd': '10',

            'T1SvodWSName': 'Страновой - Таб1',
            'T1SvodPrefix': 'T1',
            'T1SvodRowStart': '5',
            'T1SvodRowEnd': '-1',
            'T1SvodCollStart': '3',
            'T1SvodCollEnd': '21',

            'T2SvodWSName': 'Страновой - Таб2',
            'T2SvodPrefix': 'T2',
            'T2SvodRowStart': '5',
            'T2SvodRowEnd': '-1',
            'T2SvodCollStart': '2',
            'T2SvodCollEnd': '35'
        }
        # endregion

        with open(self._file_name, mode='w', encoding='utf-8') as ini_file:
            config.write(ini_file)
