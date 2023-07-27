import uuid
import xml.etree.ElementTree as ET

import argparse
from datetime import date

import openpyxl
import os
import warnings

from excel_to_list import ExcelToList
from excel_to_dict import ExcelToDict
from config_params import ConfigParams


def excel_transform_to_txt(name_of_file: str, UvedFileName: str, SvodFileName: str, el: ExcelToList):
    print("Transform to txt.\n File to parse:", name_of_file)

    warnings.simplefilter("ignore")
    workbook = openpyxl.load_workbook(name_of_file, data_only=True)
    list_d = el.uved_list_d(workbook)
    list_np = el.uved_list_np(workbook)
    list_p = el.uved_list_p(workbook)
    list_t2 = el.uved_list_t2(workbook)

    path_name = os.path.join(os.path.curdir, UvedFileName)
    with open(path_name, 'w', encoding='utf_8') as uved_file:
        uved_file.writelines(list_d)
        uved_file.writelines(list_np)
        uved_file.writelines(list_p)
        uved_file.writelines(list_t2)
    print("File ready: ", path_name)

    list_d = el.svod_list_d(workbook)
    list_np = el.svod_list_np(workbook)
    list_p = el.svod_list_p(workbook)
    list_t1 = el.svod_list_t1(workbook)
    list_t2 = el.svod_list_t2(workbook)

    path_name = os.path.join(os.path.curdir, SvodFileName)
    with open(path_name, 'w', encoding='utf_8') as svod_file:
        svod_file.writelines(list_d)
        svod_file.writelines(list_np)
        svod_file.writelines(list_p)
        svod_file.writelines(list_t1)
        svod_file.writelines(list_t2)
    print("File ready: ", path_name)

    workbook.close()
    warnings.resetwarnings()
    print("ExcelTransform to txt files finshed")


def excel_transform_to_xml(name_of_file: str, xml_file_name: str, ed: ExcelToDict):
    print("Transform to xml.\n File to parse:", name_of_file)

    warnings.simplefilter("ignore")
    workbook = openpyxl.load_workbook(name_of_file, data_only=True)
    # region Создание xml с уведомлением
    dict_d = ed.uved_dict_d(workbook)[0]
    dict_np = ed.uved_dict_np(workbook)[0]
    dict_p = ed.uved_dict_p(workbook)[0]
    list_dict_t2 = ed.uved_dict_t2(workbook)

    attr = {'ИдФайл': xml_file_name, 'ВерсФорм': dict_d['verForm'], 'ВерсПрог': dict_d['verProg']}
    root = ET.Element("Файл", attr)
    attr = {'КНД': dict_d['KND'], 'КодНО': dict_d['KodNO'], 'НомКор': dict_d['NumKor'],
            'НаимМГ': dict_d['NameMg'], 'НаимМГЛат': dict_d['NameMgLat']}
    if dict_d['NumRosMatK']:
        attr['НомРосМатК'] = dict_d['NumRosMatK']
    doc = ET.SubElement(root, "Документ", attr)
    attr = {'ДатаНач': dict_d['DateStart'], 'ДатаОконч': dict_d['DateEnd'], 'ФинГод': dict_d['FinYear']}
    ET.SubElement(doc, "ОтчПер", attr)
    attr = {'НаимУчМГЛат': dict_np['NameUchMgLat'], 'ПрУчМГ': dict_np['PrUchMg'], 'ОГРН': dict_np['OGRN'],
            'СтатУчМГ': dict_np['StatUchMg'], 'ПолнУчМГ': dict_np['PolnUchMg'], 'ПрСтрПред': dict_np['PrStrPred'],
            'ФОИВ': dict_np['FOIV'], 'ПрСогл': dict_np['PrSogl'], 'ИнфСогл': dict_np['InfSogl']}
    sub = ET.SubElement(doc, "СвНП", attr)
    attr = {'НаимОрг': dict_np['NameOrg'], 'ИННЮЛ': dict_np['INNUL'], 'КПП': dict_np['KPP']}
    ET.SubElement(sub, "НПЮЛ", attr)
    # region ---Подписант---
    attr = {'ПрПодп': dict_p['PrPodp'], 'ДолжнПодп': dict_p['DoljPodp'], 'Тлф': dict_p['Tlf'],
            'ЭлПочта': dict_p['Email']}
    sub = ET.SubElement(doc, "Подписант", attr)
    attr = {'Фамилия': dict_p['Familia'], 'Имя': dict_p['Name'], 'Отчество': dict_p['Otchestvo']}
    ET.SubElement(sub, "ФИО", attr)
    attr = {'НаимДок': dict_p['NameDok']}
    ET.SubElement(sub, "СвПред", attr)
    # endregion

    for dict_t2 in list_dict_t2:
        attr = {'СтатУчМГР2': dict_t2['StatUchMGR2'], 'ПрУчМГ': dict_t2['PrUchMG'],
                'ИНН': dict_t2['INN'], 'КПП': dict_t2['KPP'], 'НаимУч': dict_t2['NameUch'],
                'НаимУчЛат': dict_t2['NameUchLat'], 'ПрСтрПред': dict_t2['PrStrPred']}
        if dict_t2['StatUchMGR2'] == '2':
            attr['ОсновПредстСтран'] = dict_t2['OsnovPredStran']
        if dict_t2['PrUchMG'] == '1':
            attr['ОГРН'] = dict_t2['OGRN']
        if dict_t2['PrStrPred'] != '1':
            attr['ФОИВ'] = dict_t2['FOIV']
        if dict_t2['PrSogl']:
            attr['ПрСогл'] = dict_t2['PrSogl']
        if dict_t2['InfSogl']:
            attr['ИнфСогл'] = dict_t2['InfSogl']
        ET.SubElement(doc, "Раздел2", attr)

    # save to file
    path_name = os.path.join(os.path.curdir, xml_file_name)
    tree = ET.ElementTree(root)
    tree.write(path_name + ".xml", encoding="windows-1251")
    print("File ready: ", path_name + ".xml")
    # endregion

    # region Создание свода
    dict_d = ed.svod_dict_d(workbook)[0]
    dict_np = ed.svod_dict_np(workbook)[0]
    dict_p = ed.svod_dict_p(workbook)[0]
    list_dict_t1 = ed.svod_dict_t1(workbook)
    list_dict_t2 = ed.svod_dict_t2(workbook)

    workbook.close()
    warnings.resetwarnings()

    xml_file_name = gen_file_name("ON_STRANOTCH")
    attr = {'ИдФайл': xml_file_name, 'ВерсФорм': dict_d['VerForm'], 'ВерсПрог': dict_d['VerProg']}
    root = ET.Element("Файл", attr)
    attr = {'КНД': dict_d['KND'], 'КодНО': dict_d['KodNO'], 'НомКор': dict_d['NumKor'],
            'ПринСтр': dict_d['PrinCnt'], 'Язык': dict_d['Language'], 'НаимМГ': dict_d['NameMG'],
            'НаимМГЛат': dict_d['NameMGLa'], 'НомРосМатК': dict_d['NumMatK'],
            'ДатаВрФорм': str(dict_d['DateVrForm']).replace(' ', 'T'),
            'ПредупрИнф': dict_d['PreduprInf']}
    doc = ET.SubElement(root, "Документ", attr)
    attr = {'ДатаНач': dict_d['DateNach'], 'ДатаОконч': dict_d['DateOkonchan'], 'ФинГод': dict_d['FinYear']}
    ET.SubElement(doc, "ОтчПер", attr)
    country_np = dict_d['CountryNaprOtch']
    for item in country_np.split(', '):
        sub = ET.SubElement(doc, "СтранНапрОтч")
        sub.text = item
    sub = ET.SubElement(doc, "СвНП")
    attr = {}
    attr['НаимОрг'] = dict_np['NameOrg']
    attr['ИННЮЛ'] = dict_np['INNUL']
    attr['КПП'] = dict_np['KPP']
    ET.SubElement(sub, "НПЮЛ", attr)

    #  ---Раздел1---
    # region ---наполнение атрибутов Раздел1---
    attr = {}
    attr['НомКорРазд1'] = dict_np['NumKorRazd1']
    attr['ИдРазд1'] = str(uuid.uuid4())
    attr['СтатУчМГ'] = dict_np['StatUchMG']
    attr['СтрНалРезид'] = dict_np['StrNalRezid']
    attr['ПрУчМГ'] = dict_np['PrUchMg']
    attr['СтрНомНП'] = dict_np['StrNumNP']
    if dict_np['RegNom']:
        attr['РегНом'] = dict_np['RegNom']
    if dict_np['TypeRegNom']:
        attr['ТипРегНом'] = dict_np['TypeRegNom']
    if dict_np['TypeRegNomLat']:
        attr['ТипРегНомЛат'] = dict_np['TypeRegNomLat']
    if dict_np['StrRegNom']:
        attr['СтрРегНом'] = dict_np['StrRegNom']
    attr['НаимУчЛат'] = dict_np['NameUchLat']
    # endregion
    razd1 = ET.SubElement(sub, "Раздел1", attr)
    # region ---наполнение атрибутов узлов СвАдрес, Адрес---
    attr = {}
    if dict_np['AddrInoyLat']:
        attr['АдрИнойЛат'] = dict_np['AddrInoyLat']
    if dict_np['AddrInoy']:
        attr['АдрИной'] = dict_np['AddrInoy']
    attr['ТипАдрес'] = dict_np['TypeAddr']
    attr['СтрАдр'] = dict_np['StrAddr']
    svAddr = ET.SubElement(razd1, "СвАдрес", attr)
    attr = {}
    if dict_np['AbYashik']:
        attr['АбЯщик'] = '{0}______'.format(dict_np['AbYashik'])[:6]
    if dict_np['Appartament']:
        attr['Квартира'] = dict_np['Appartament']
    if dict_np['Korpus']:
        attr['Корпус'] = dict_np['Korpus']
    if dict_np['Home']:
        attr['Дом'] = dict_np['Home']
    if dict_np['StreetLat']:
        attr['УлицаЛат'] = dict_np['StreetLat']
    if dict_np['Street']:
        attr['Улица'] = dict_np['Street']
    if dict_np['RaionLat']:
        attr['РайонЛат'] = dict_np['RaionLat']
    if dict_np['Raion']:
        attr['Район'] = dict_np['Raion']
    attr['ГородЛат'] = dict_np['CityLat']
    attr['Город'] = dict_np['City']
    if dict_np['RegionLat']:
        attr['РегионЛат'] = dict_np['RegionLat']
    if dict_np['Region']:
        attr['Регион'] = dict_np['Region']
    if dict_np['IndexIn']:
        attr['ИндексИн'] = dict_np['IndexIn']
    if dict_np['Index']:
        attr['ИндексРос'] = dict_np['Index']
    # endregion
    ET.SubElement(svAddr, "Адрес", attr)
    # region ---Заполнение узлов Подписант, ФИО, СвПред---
    attr = {'ПрПодп': dict_p['PrPodp'], 'ДолжнПодп': dict_p['DoljnPod'], 'КонтИнф': dict_p['KontInf'],
            'КонтИнфЛат': dict_p['KontInfLat']}
    sub = ET.SubElement(doc, "Подписант", attr)
    attr = {'Фамилия': dict_p['Familiya'], 'Имя': dict_p['Name'], 'Отчество': dict_p['Otchestvo']}
    ET.SubElement(sub, "ФИО", attr)
    attr = {'НаимДок': dict_p['NameDok']}
    ET.SubElement(sub, "СвПред", attr)
    # endregion

    # ---Раздел2---
    for dict_t1 in list_dict_t1:
        idRazd2 = dict_t1['IdRazd2'] if dict_t1['NomKorRazd2'] == '999' else str(uuid.uuid4())
        attr = {'НомКорРазд2': dict_t1['NomKorRazd2'], 'ИдРазд2': idRazd2, 'СтрНалРезид': dict_t1['StrNalRezid']}
        razdel2 = ET.SubElement(doc, "Раздел2", attr)
        # region ---заполнение раздела ПокДеятУчМГ---
        idPokD = str(uuid.uuid4())
        if (dict_t1['NomKorRazd2'] == '0' or dict_t1['NomKorRazd2'] == '999') and (dict_d['NumKor'] != '0'):
            idPokD = dict_t1['IdPokDeyatUchMG']

        attr = {'ИдПокДеятУчМГ': idPokD, 'ЧислРабШтат': dict_t1['ChislRab'],
                'ЧислРабНез': dict_t1['ChislRabNez'], 'ЧислРабВсего': dict_t1['ChislRabVsego']}
        if dict_t1['DopInfChislRab']:
            attr['ДопИнфЧислРаб'] = dict_t1['DopInfChislRab']
        sub = ET.SubElement(razdel2, "ПокДеятУчМГ", attr)
        attr = {'КодВал': 'RUB', 'Сум': dict_t1['DohUchMG']}
        ET.SubElement(sub, "ДохУчМГ", attr)
        attr = {'КодВал': 'RUB', 'Сум': dict_t1['DohNezLic']}
        ET.SubElement(sub, "ДохНезЛиц", attr)
        attr = {'КодВал': 'RUB', 'Сум': dict_t1['DohVsego']}
        ET.SubElement(sub, "ДохВсего", attr)
        attr = {'КодВал': 'RUB', 'Сум': dict_t1['PribDoNal']}
        ET.SubElement(sub, "ПрибДоНал", attr)
        attr = {'КодВал': 'RUB', 'Сум': dict_t1['NalPribUpl']}
        ET.SubElement(sub, "НалПрибУпл", attr)
        attr = {'КодВал': 'RUB', 'Сум': dict_t1['NalPribNach']}
        ET.SubElement(sub, "НалПрибНач", attr)
        attr = {'КодВал': 'RUB', 'Сум': dict_t1['Capital']}
        ET.SubElement(sub, "Капитал", attr)
        dobcap = dict_t1['DobCapital']
        if dobcap:
            attr = {'КодВал': 'RUB', 'Сум': dobcap}
            ET.SubElement(sub, "ДобКапитал", attr)
        attr = {'КодВал': 'RUB', 'Сум': dict_t1['NakPrib']}
        ET.SubElement(sub, "НакПриб", attr)
        attr = {'КодВал': 'RUB', 'Сум': dict_t1['Activ']}
        ET.SubElement(sub, "Активы", attr)
        # endregion

        # ---заполнение раздела УчастникМГ---
        for dict_t2 in list_dict_t2:
            if dict_t2['StrNalRezid'] == dict_t1['StrNalRezid']:
                # region ---заполнение атрибутов УчастникМГ---
                attr = {}
                if dict_t1['StrNalRezid'] != 'RU':
                    attr['НомНП'] = dict_t2['NomNP']
                elif dict_t2['PrUchMg'] == '3' and dict_t2['NomNP']:
                    attr['НомНП'] = dict_t2['NomNP']
                attr['НаимУч'] = dict_t2['NameUch']
                attr['СтатУчМГ'] = dict_t2['StatUchMg']
                attr['ПрУчМГ'] = dict_t2['PrUchMg']
                attr['СтрНалРезид'] = dict_t2['StrNalRezid']
                attr['СтрНомНП'] = dict_t2['StrNomNP']
                if dict_t2['RegNom']:  # dict_t2['StrRegNom'] == 'RU' and
                    attr['РегНом'] = dict_t2['RegNom']
                if dict_t2['TypeRegNom']:
                    attr['ТипРегНом'] = dict_t2['TypeRegNom']
                if dict_t2['TypeRegNomLat']:
                    attr['ТипРегНомЛат'] = dict_t2['TypeRegNomLat']
                if dict_t2['StrRegNom']:
                    attr['СтрРегНом'] = dict_t2['StrRegNom']
                attr['НаимУчЛат'] = dict_t2['NameUchLat']
                if dict_t1['StrNalRezid'] == 'RU':
                    attr['ИННЮЛ'] = dict_t2['INNUL']
                elif dict_t2['PrUchMg'] == '3' and dict_t2['INNUL']:
                    attr['ИННЮЛ'] = dict_t2['INNUL']
                # endregion
                sub = ET.SubElement(razdel2, "УчастникМГ", attr)

                # region ---заполнение атрибутов СвАдрес---
                attr = {}
                if dict_t2['AddrInoy']:
                    attr['АдрИной'] = dict_t2['AddrInoy']
                if dict_t2['AddrInoyLat']:
                    attr['АдрИнойЛат'] = dict_t2['AddrInoyLat']
                attr['ТипАдрес'] = dict_t2['TypeAddr']
                attr['СтрАдр'] = dict_t2['StrAddr']
                # endregion
                svAddr = ET.SubElement(sub, "СвАдрес", attr)
                # region ---наполнение значений атрибутов узла Адрес---
                attr = {}
                if dict_t2['AbYashik']:
                    attr['АбЯщик'] = '{0}______'.format(dict_t2['AbYashik'])[:6]
                if dict_t2['Appartament']:
                    attr['Квартира'] = dict_t2['Appartament']
                if dict_t2['Korpus']:
                    attr['Корпус'] = dict_t2['Korpus']
                if dict_t2['Home']:
                    attr['Дом'] = dict_t2['Home']
                if dict_t2['StreetLat']:
                    attr['УлицаЛат'] = dict_t2['StreetLat']
                if dict_t2['Street']:
                    attr['Улица'] = dict_t2['Street']
                if dict_t2['RaionLat']:
                    attr['РайонЛат'] = dict_t2['RaionLat']
                if dict_t2['Raion']:
                    attr['Район'] = dict_t2['Raion']
                attr['ГородЛат'] = dict_t2['CityLat']
                attr['Город'] = dict_t2['City']
                if dict_t2['RegionLat']:
                    attr['РегионЛат'] = dict_t2['RegionLat']
                if dict_t2['Region']:
                    attr['Регион'] = dict_t2['Region']
                if dict_t2['StrAddr'] != 'RU' and dict_t2['IndexIn']:
                    attr['ИндексИн'] = dict_t2['IndexIn']
                if dict_t2['IndexRos']:
                    attr['ИндексРос'] = dict_t2['IndexRos']
                # endregion
                ET.SubElement(svAddr, "Адрес", attr)
                attr = {}
                if dict_t2['DopInfUch']:
                    attr['ДопИнфУч'] = dict_t2['DopInfUch']
                attr['СтрРегИнк'] = dict_t2['StrRegInk']
                typDop = ET.SubElement(sub, "СвСтранТипДоп", attr)
                typEk = dict_t2['TypeEkDeyat']
                for item in typEk.split(','):
                    if item:
                        sub = ET.SubElement(typDop, 'ТипЭкДеят')
                        sub.text = item

    # save to file
    path_name = os.path.join(os.path.curdir, xml_file_name)
    tree = ET.ElementTree(root)
    tree.write(path_name + ".xml", encoding="windows-1251")
    print("File ready: ", path_name + ".xml")
    # endregion

    print("ExcelTransform to xml files finshed")

def gen_file_name(prm: str) -> str:
    strtoday = date.today().strftime('%Y%m%d')
    id = str(uuid.uuid4()).replace('-', '')
    return prm + "_9972_9972_7736050003997250001_" + strtoday + "_" + id

def parse_excel():
    print("ExcelTransform started")
    parser = argparse.ArgumentParser()
    parser.add_argument('filename', help='name of excel file with data')
    parser.add_argument('-xml', action='store_true', help='under construction testing to create xml')
    parser.add_argument('-ini', action='store_false', help='creates ini-file with offsets in excel file')

    args = parser.parse_args()
    cfg = ConfigParams()
    if not args.ini:
        name_of_file = args.filename
        if not args.xml:
            excel_transform_to_txt(name_of_file, cfg.UvedFileName, cfg.SvodFileName, el = ExcelToList(cfg))
        else:
            excel_transform_to_xml(name_of_file, xml_file_name = gen_file_name("ON_UVUCHMGR"), ed = ExcelToDict(cfg))
    else:
        cfg.create_ini_file()
