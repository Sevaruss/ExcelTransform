from config_params import ConfigParams


class ExcelToDict:
    def __init__(self, cfg: ConfigParams):
        self.cfg: ConfigParams = cfg

    def replace_invalid(self, val: str) -> str:
        val = '' if val is None else str(val)
        val = val.replace('\n', '').replace('\r', '').replace(' 00:00:00', '')
        return val

    field_uved_d = ['KEY', 'idFile', 'verForm', 'verProg', 'KND', 'KodNO', 'NumKor', 'NameMg', 'NameMgLat',
                    'NumRosMatK', 'DateStart', 'DateEnd', 'FinYear']
    field_uved_np = ['KEY', 'NameOrg', 'INNUL', 'KPP', 'NameUchMgLat', 'PrUchMg', 'OGRN', 'StatUchMg', 'PolnUchMg',
                     'OsnovPredstStran', 'PrStrPred', 'FOIV', 'PrSogl', 'InfSogl']
    field_uved_p = ['KEY', 'PrPodp', 'DoljPodp', 'Tlf', 'Email', 'Familia', 'Name', 'Otchestvo', 'NameDok']
    field_uved_t2 = ['KEY', 'StatUchMGR2', 'OsnovPredStran', 'PrUchMG', 'OGRN', 'INN', 'KPP', 'NameUch', 'NameUchLat',
                     'OsnovPredstUved', 'PrStrPred', 'FOIV', 'PrSogl', 'InfSogl']
    field_svod_d = ['KEY', 'IdFile', 'VerForm', 'VerProg', 'KND', 'KodNO', 'NumKor', 'PrinCnt', 'Language', 'NameMG',
                    'NameMGLa', 'NumMatK', 'DateVrForm', 'PreduprInf', 'DateNach', 'DateOkonchan', 'FinYear',
                    'CountryNaprOtch']
    field_svod_np = ['KEY', 'NameOrg', 'INNUL', 'KPP', 'NumKorRazd1', 'IdRazd1', 'StatUchMG', 'PrUchMg', 'StrNalRezid',
                     'StrNumNP', 'RegNom', 'TypeRegNom', 'TypeRegNomLat', 'StrRegNom', 'NameUchLat', 'TypeAddr',
                     'StrAddr', 'AddrInoy', 'AddrInoyLat', 'Index', 'IndexIn', 'AbYashik', 'Region', 'RegionLat',
                     'Raion', 'RaionLat', 'City', 'CityLat', 'Street', 'StreetLat', 'Home', 'Korpus', 'Appartament']
    field_svod_p = ['KEY', 'PrPodp', 'DoljnPod', 'KontInf', 'KontInfLat', 'Familiya', 'Name', 'Otchestvo', 'NameDok']
    field_svod_t1 = ['KEY', 'StrNalRezid', 'NomKorRazd2', 'IdRazd2', 'IdPokDeyatUchMG', 'DohVsego', 'DohUchMG',
                     'DohNezLic', 'PribDoNal', 'NalPribUpl', 'NalPribNach', 'Capital', 'DobCapital', 'NakPrib', 'Activ',
                     'ChislRab', 'ChislRabNez', 'ChislRabVsego', 'DopInfChislRab']
    field_svod_t2 = ['KEY', 'StatUchMg', 'PrUchMg', 'StrNalRezid', 'NomNP', 'INNUL', 'StrNomNP', 'RegNom', 'TypeRegNom',
                     'TypeRegNomLat', 'StrRegNom', 'NameUch', 'NameUchLat', 'TypeAddr', 'StrAddr', 'AddrInoy',
                     'AddrInoyLat', 'IndexRos', 'IndexIn', 'AbYashik', 'Region', 'RegionLat', 'Raion', 'RaionLat',
                     'City', 'CityLat', 'Street', 'StreetLat', 'Home', 'Korpus', 'Appartament', 'StrRegInk',
                     'DopInfUch', 'TypeEkDeyat']

    def rows_dict(self, workbook, ws_name: str, field_names, row_prefix: str, coll_start: int, coll_end: int,
                  row_start: int, row_end: int) -> list:
        ws = workbook[ws_name]
        big_list = []
        i = row_start
        end_row = ws.max_row if row_end == -1 else row_end
        while i <= end_row:
            fn = 1
            one_row = {field_names[0]: row_prefix}
            for cnt in range(coll_start, coll_end):
                val = self.replace_invalid(ws.cell(row=i, column=cnt).value)
                other = {field_names[fn]: val}
                one_row.update(other)
                fn = fn + 1
            big_list = big_list + [one_row]
            i = i + 1
        return big_list

    def uved_dict_d(self, workbook) -> list:
        return self.rows_dict(workbook, self.cfg.DUvedWSName, self.field_uved_d, self.cfg.DUvedPrefix,
                              self.cfg.DUvedCollStart, self.cfg.DUvedCollEnd, self.cfg.DUvedRowStart,
                              self.cfg.DUvedRowEnd)

    def uved_dict_np(self, workbook) -> list:
        return self.rows_dict(workbook, self.cfg.NPUvedWSName, self.field_uved_np, self.cfg.NPUvedPrefix,
                              self.cfg.NPUvedCollStart, self.cfg.NPUvedCollEnd, self.cfg.NPUvedRowStart,
                              self.cfg.NPUvedRowEnd)

    def uved_dict_p(self, workbook) -> list:
        return self.rows_dict(workbook, self.cfg.PUvedWSName, self.field_uved_p, self.cfg.PUvedPrefix,
                              self.cfg.PUvedCollStart, self.cfg.PUvedCollEnd, self.cfg.PUvedRowStart,
                              self.cfg.PUvedRowEnd)

    def uved_dict_t2(self, workbook) -> list:
        return self.rows_dict(workbook, self.cfg.T2UvedWSName, self.field_uved_t2, self.cfg.T2UvedPrefix,
                              self.cfg.T2UvedCollStart, self.cfg.T2UvedCollEnd, self.cfg.T2UvedRowStart,
                              self.cfg.T2UvedRowEnd)

    def svod_dict_d(self, workbook) -> list:
        return self.rows_dict(workbook, self.cfg.DSvodWSName, self.field_svod_d, self.cfg.DSvodPrefix,
                              self.cfg.DSvodCollStart, self.cfg.DSvodCollEnd, self.cfg.DSvodRowStart,
                              self.cfg.DSvodRowEnd)

    def svod_dict_np(self, workbook) -> list:
        return self.rows_dict(workbook, self.cfg.NPSvodWSName, self.field_svod_np, self.cfg.NPSvodPrefix,
                              self.cfg.NPSvodCollStart, self.cfg.NPSvodCollEnd, self.cfg.NPSvodRowStart,
                              self.cfg.NPSvodRowEnd)

    def svod_dict_p(self, workbook) -> list:
        return self.rows_dict(workbook, self.cfg.PSvodWSName, self.field_svod_p, self.cfg.PSvodPrefix,
                              self.cfg.PSvodCollStart, self.cfg.PSvodCollEnd, self.cfg.PSvodRowStart,
                              self.cfg.PSvodRowEnd)

    def svod_dict_t1(self, workbook) -> list:
        return self.rows_dict(workbook, self.cfg.T1SvodWSName, self.field_svod_t1, self.cfg.T1SvodPrefix,
                              self.cfg.T1SvodCollStart, self.cfg.T1SvodCollEnd, self.cfg.T1SvodRowStart,
                              self.cfg.T1SvodRowEnd)

    def svod_dict_t2(self, workbook) -> list:
        return self.rows_dict(workbook, self.cfg.T2SvodWSName, self.field_svod_t2, self.cfg.T2SvodPrefix,
                              self.cfg.T2SvodCollStart, self.cfg.T2SvodCollEnd, self.cfg.T2SvodRowStart,
                              self.cfg.T2SvodRowEnd)
