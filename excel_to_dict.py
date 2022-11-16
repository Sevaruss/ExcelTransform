import config_params as cfg


def replace_invalid(val: str) -> str:
    val = '' if val is None else str(val)
    val = val.replace('\n', '').replace('\r', '').replace(' 00:00:00', '')
    return val


field_uved_d = ['KEY', 'idFile', 'verForm', 'verProg', 'KND', 'KodNO', 'NumKor', 'NameMg', 'NameMgLat', 'NumRosMatK',
                'DateStart', 'DateEnd', 'FinYear']
field_uved_np = ['KEY', 'NameOrg', 'INNUL', 'KPP', 'NameUchMgLat', 'PrUchMg', 'OGRN', 'StatUchMg', 'PolnUchMg',
                 'OsnovPredstStran', 'PrStrPred', 'FOIV', 'PrSogl', 'InfSogl']
field_uved_p = ['KEY', 'PrPodp', 'DoljPodp', 'Tlf', 'Email', 'Familia', 'Name', 'Otchestvo', 'NameDok']
field_uved_t2 = ['KEY', 'StatUchMGR2', 'OsnovPredStran', 'PrUchMG', 'OGRN', 'INN', 'KPP', 'NameUch', 'NameUchLat',
                 'OsnovPredstUved', 'PrStrPred', 'FOIV', 'PrSogl', 'InfSogl']
field_svod_d = ['KEY', 'IdFile', 'VerForm', 'VerProg', 'KND', 'KodNO', 'NumKor', 'PrinCnt', 'Language', 'NameMG',
                'NameMGLa', 'NumMatK', 'DateVrForm', 'PreduprInf', 'DateNach', 'DateOkonchan', 'FinYear',
                'CountryNaprOtch']
field_svod_np = ['KEY', 'NameOrg', 'INNUL', 'KPP', 'NumKorRazd1', 'IdRazd1', 'StatUchMG', 'PrUchMg', 'StrNalRezid',
                 'StrNumNP', 'RegNom', 'TypeRegNom', 'TypeRegNomLat', 'StrRegNom', 'NameUchLat', 'TypeAddr', 'StrAddr',
                 'AddrInoy', 'AddrInoyLat', 'Index', 'IndexIn', 'AbYashik', 'Region', 'RegionLat', 'Raion', 'RaionLat',
                 'City', 'CityLat', 'Street', 'StreetLat', 'Home', 'Korpus', 'Appartament']
field_svod_p = ['KEY', 'PrPodp', 'DoljnPod', 'KontInf', 'KontInfLat', 'Familiya', 'Name', 'Otchestvo', 'NameDok']
field_svod_t1 = ['KEY', 'StrNalRezid', 'NomKorRazd2', 'IdRazd2', 'IdPokDeyatUchMG', 'DohVsego', 'DohUchMG', 'DohNezLic',
                 'PribDoNal', 'NalPribUpl', 'NalPribNach', 'Capital', 'DobCapital', 'NakPrib', 'Activ', 'ChislRab',
                 'ChislRabNez', 'ChislRabVsego', 'DopInfChislRab']
field_svod_t2 = ['KEY', 'StatUchMg', 'PrUchMg', 'StrNalRezid', 'NomNP', 'INNUL', 'StrNomNP', 'RegNom', 'TypeRegNom',
                 'TypeRegNomLat', 'StrRegNom', 'NameUch', 'NameUchLat', 'TypeAddr', 'StrAddr', 'AddrInoy',
                 'AddrInoyLat', 'IndexRos', 'IndexIn', 'AbYashik', 'Region', 'RegionLat', 'Raion', 'RaionLat', 'City',
                 'CityLat', 'Street', 'StreetLat', 'Home', 'Korpus', 'Appartament', 'StrRegInk', 'DopInfUch',
                 'TypeEkDeyat']


def rows_dict(workbook, ws_name: str, field_names, row_prefix: str, coll_start: int, coll_end: int, row_start: int,
              row_end: int) -> list:
    ws = workbook[ws_name]
    big_list = []
    i = row_start
    end_row = ws.max_row if row_end == -1 else row_end
    while i <= end_row:
        fn = 1
        one_row = {field_names[0]: row_prefix}
        for cnt in range(coll_start, coll_end):
            val = replace_invalid(ws.cell(row=i, column=cnt).value)
            other = {field_names[fn]: val}
            one_row.update(other)
            fn = fn + 1
        big_list = big_list + [one_row]
        i = i + 1
    return big_list


def uved_dict_d(workbook) -> list:
    return rows_dict(workbook, cfg.DUvedWSName, field_uved_d, cfg.DUvedPrefix, cfg.DUvedCollStart, cfg.DUvedCollEnd,
                     cfg.DUvedRowStart, cfg.DUvedRowEnd)


def uved_dict_np(workbook) -> list:
    return rows_dict(workbook, cfg.NPUvedWSName, field_uved_np, cfg.NPUvedPrefix, cfg.NPUvedCollStart,
                     cfg.NPUvedCollEnd, cfg.NPUvedRowStart, cfg.NPUvedRowEnd)


def uved_dict_p(workbook) -> list:
    return rows_dict(workbook, cfg.PUvedWSName, field_uved_p, cfg.PUvedPrefix, cfg.PUvedCollStart, cfg.PUvedCollEnd,
                     cfg.PUvedRowStart, cfg.PUvedRowEnd)


def uved_dict_t2(workbook) -> list:
    return rows_dict(workbook, cfg.T2UvedWSName, field_uved_t2, cfg.T2UvedPrefix, cfg.T2UvedCollStart,
                     cfg.T2UvedCollEnd, cfg.T2UvedRowStart, cfg.T2UvedRowEnd)


def svod_dict_d(workbook) -> list:
    return rows_dict(workbook, cfg.DSvodWSName, field_svod_d, cfg.DSvodPrefix, cfg.DSvodCollStart, cfg.DSvodCollEnd,
                     cfg.DSvodRowStart, cfg.DSvodRowEnd)


def svod_dict_np(workbook) -> list:
    return rows_dict(workbook, cfg.NPSvodWSName, field_svod_np, cfg.NPSvodPrefix, cfg.NPSvodCollStart,
                     cfg.NPSvodCollEnd, cfg.NPSvodRowStart, cfg.NPSvodRowEnd)


def svod_dict_p(workbook) -> list:
    return rows_dict(workbook, cfg.PSvodWSName, field_svod_p, cfg.PSvodPrefix, cfg.PSvodCollStart, cfg.PSvodCollEnd,
                     cfg.PSvodRowStart, cfg.PSvodRowEnd)


def svod_dict_t1(workbook) -> list:
    return rows_dict(workbook, cfg.T1SvodWSName, field_svod_t1, cfg.T1SvodPrefix, cfg.T1SvodCollStart,
                     cfg.T1SvodCollEnd, cfg.T1SvodRowStart, cfg.T1SvodRowEnd)


def svod_dict_t2(workbook) -> list:
    return rows_dict(workbook, cfg.T2SvodWSName, field_svod_t2, cfg.T2SvodPrefix, cfg.T2SvodCollStart,
                     cfg.T2SvodCollEnd, cfg.T2SvodRowStart, cfg.T2SvodRowEnd)
