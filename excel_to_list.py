from config_params import ConfigParams


class ExcelToList:

    def __init__(self, cfg: ConfigParams):
        self.cfg: ConfigParams = cfg

    def replace_invalids(self, val: str):
        val = ';' if val is None else str(val) + ';'
        val = val.replace('\n', '').replace('\r', '').replace(' 00:00:00', '')
        return val

    def rows_list(self, workbook, ws_name: str, row_prefix: str, coll_start: int, coll_end: int,
                  row_start: int, row_end: int) -> list:
        ws = workbook[ws_name]
        big_list = []
        i = row_start
        end_row = ws.max_row if row_end == -1 else row_end
        while i <= end_row:
            one_row = [row_prefix + ';']
            for cnt in range(coll_start, coll_end):
                one_row = one_row + [self.replace_invalids(ws.cell(row=i, column=cnt).value)]
            big_list = big_list + [''.join(one_row) + '\n']
            i = i + 1
        return big_list

    def uved_list_d(self, workbook) -> list:
        return self.rows_list(workbook, self.cfg.DUvedWSName, self.cfg.DUvedPrefix, self.cfg.DUvedCollStart,
                              self.cfg.DUvedCollEnd, self.cfg.DUvedRowStart, self.cfg.DUvedRowEnd)

    def uved_list_np(self, workbook) -> list:
        return self.rows_list(workbook, self.cfg.NPUvedWSName, self.cfg.NPUvedPrefix, self.cfg.NPUvedCollStart,
                              self.cfg.NPUvedCollEnd, self.cfg.NPUvedRowStart, self.cfg.NPUvedRowEnd)

    def uved_list_p(self, workbook) -> list:
        return self.rows_list(workbook, self.cfg.PUvedWSName, self.cfg.PUvedPrefix, self.cfg.PUvedCollStart,
                              self.cfg.PUvedCollEnd, self.cfg.PUvedRowStart, self.cfg.PUvedRowEnd)

    def uved_list_t2(self, workbook) -> list:
        return self.rows_list(workbook, self.cfg.T2UvedWSName, self.cfg.T2UvedPrefix, self.cfg.T2UvedCollStart,
                              self.cfg.T2UvedCollEnd, self.cfg.T2UvedRowStart, self.cfg.T2UvedRowEnd)

    def svod_list_d(self, workbook) -> list:
        return self.rows_list(workbook, self.cfg.DSvodWSName, self.cfg.DSvodPrefix, self.cfg.DSvodCollStart,
                              self.cfg.DSvodCollEnd, self.cfg.DSvodRowStart, self.cfg.DSvodRowEnd)

    def svod_list_np(self, workbook) -> list:
        return self.rows_list(workbook, self.cfg.NPSvodWSName, self.cfg.NPSvodPrefix, self.cfg.NPSvodCollStart,
                              self.cfg.NPSvodCollEnd, self.cfg.NPSvodRowStart, self.cfg.NPSvodRowEnd)

    def svod_list_p(self, workbook) -> list:
        return self.rows_list(workbook, self.cfg.PSvodWSName, self.cfg.PSvodPrefix, self.cfg.PSvodCollStart,
                              self.cfg.PSvodCollEnd, self.cfg.PSvodRowStart, self.cfg.PSvodRowEnd)

    def svod_list_t1(self, workbook) -> list:
        return self.rows_list(workbook, self.cfg.T1SvodWSName, self.cfg.T1SvodPrefix, self.cfg.T1SvodCollStart,
                              self.cfg.T1SvodCollEnd, self.cfg.T1SvodRowStart, self.cfg.T1SvodRowEnd)

    def svod_list_t2(self, workbook) -> list:
        return self.rows_list(workbook, self.cfg.T2SvodWSName, self.cfg.T2SvodPrefix, self.cfg.T2SvodCollStart,
                              self.cfg.T2SvodCollEnd, self.cfg.T2SvodRowStart, self.cfg.T2SvodRowEnd)

# def test_list(workbook, wsname):
#     ws = workbook[wsname]
#     for row in ws.rows:
#         print('row: !!!!!!!!!!!!!!!!!!!!!!!!')
#         for cell in row:
#             print(cell, cell.value)
