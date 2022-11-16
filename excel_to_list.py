import config_params as cfg


def replace_invalids(val: str):
    val = ';' if val is None else str(val) + ';'
    val = val.replace('\n', '').replace('\r', '').replace(' 00:00:00', '')
    return val


def rows_list(workbook, ws_name: str, row_prefix: str, coll_start: int, coll_end: int,
              row_start: int, row_end: int) -> list:
    ws = workbook[ws_name]
    big_list = []
    i = row_start
    end_row = ws.max_row if row_end == -1 else row_end
    while i <= end_row:
        one_row = [row_prefix + ';']
        for cnt in range(coll_start, coll_end):
            one_row = one_row + [replace_invalids(ws.cell(row=i, column=cnt).value)]
        big_list = big_list + [''.join(one_row) + '\n']
        i = i + 1
    return big_list


def uved_list_d(workbook) -> list:
    return rows_list(workbook, cfg.DUvedWSName, cfg.DUvedPrefix, cfg.DUvedCollStart, cfg.DUvedCollEnd,
                     cfg.DUvedRowStart, cfg.DUvedRowEnd)


def uved_list_np(workbook) -> list:
    return rows_list(workbook, cfg.NPUvedWSName, cfg.NPUvedPrefix, cfg.NPUvedCollStart, cfg.NPUvedCollEnd,
                     cfg.NPUvedRowStart, cfg.NPUvedRowEnd)


def uved_list_p(workbook) -> list:
    return rows_list(workbook, cfg.PUvedWSName, cfg.PUvedPrefix, cfg.PUvedCollStart, cfg.PUvedCollEnd,
                     cfg.PUvedRowStart, cfg.PUvedRowEnd)


def uved_list_t2(workbook) -> list:
    return rows_list(workbook, cfg.T2UvedWSName, cfg.T2UvedPrefix, cfg.T2UvedCollStart, cfg.T2UvedCollEnd,
                     cfg.T2UvedRowStart, cfg.T2UvedRowEnd)


def svod_list_d(workbook) -> list:
    return rows_list(workbook, cfg.DSvodWSName, cfg.DSvodPrefix, cfg.DSvodCollStart, cfg.DSvodCollEnd,
                     cfg.DSvodRowStart, cfg.DSvodRowEnd)


def svod_list_np(workbook) -> list:
    return rows_list(workbook, cfg.NPSvodWSName, cfg.NPSvodPrefix, cfg.NPSvodCollStart, cfg.NPSvodCollEnd,
                     cfg.NPSvodRowStart, cfg.NPSvodRowEnd)


def svod_list_p(workbook) -> list:
    return rows_list(workbook, cfg.PSvodWSName, cfg.PSvodPrefix, cfg.PSvodCollStart, cfg.PSvodCollEnd,
                     cfg.PSvodRowStart, cfg.PSvodRowEnd)


def svod_list_t1(workbook) -> list:
    return rows_list(workbook, cfg.T1SvodWSName, cfg.T1SvodPrefix, cfg.T1SvodCollStart, cfg.T1SvodCollEnd,
                     cfg.T1SvodRowStart, cfg.T1SvodRowEnd)


def svod_list_t2(workbook) -> list:
    return rows_list(workbook, cfg.T2SvodWSName, cfg.T2SvodPrefix, cfg.T2SvodCollStart, cfg.T2SvodCollEnd,
                     cfg.T2SvodRowStart, cfg.T2SvodRowEnd)


# def test_list(workbook, wsname):
#     ws = workbook[wsname]
#     for row in ws.rows:
#         print('row: !!!!!!!!!!!!!!!!!!!!!!!!')
#         for cell in row:
#             print(cell, cell.value)