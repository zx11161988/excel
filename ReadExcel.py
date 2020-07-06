#! /user/bin/env python
from openpyxl import load_workbook


class ReadData:
    array = ["地铁线", "片区", "小区名称", "户型", "方位", "面积", "总价", "均价", "楼层", "装修", "建造时间", "区域", "挂牌时间", "url"]
    info = {"地铁线": "", "片区": "", "小区名称": "", "户型": "", "方位": "", "面积": "", "总价": "",
            "均价": "", "楼层": "", "装修": "", "建造时间": "", "区域": "", "挂牌时间": "", "url": ""}

    number = 1
    path = "./2020年上海三零卫士月度决算表-成都6 - 定稿.xlsx"

    # wb = openpyxl.Workbook()
    # sheet = wb.active
    # sheet.title = '成都合同明细'

    def __init__(self, tile_array):
        workbook1 = load_workbook('2020年上海三零卫士月度决算表-成都6 - 定稿.xlsx')
        sheet = workbook1['成都合同明细']

    def _readData(self, array, sheet=None):
        workbook1 = load_workbook('D:/StudyByDoing/fangjia_lianjia-master/excel/2020年上海三零卫士月度决算表-成都6 - 定稿.xlsx')
        sheet = workbook1['成都合同明细']
        max_row = sheet.max_row
        data = []
        for row in sheet.iter_rows(min_row=1, max_col=3, max_row=2):
            for cell in row:
                print(cell)
        #print("读取到的所有测试用例：", data)
        #print("读取数据成功！")

    def read(self, dict_value=None):
        print("read >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>")
        # info_array = []
        # for item in WriteData.array:
        #    print("write", item)
        #    info_array.append(dict_value.get(item))
        self._readData(dict_value)
        print("read <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<")
        # _writeData(info_array)


if __name__ == '__main__':
    data = ReadData(ReadData.array)
    data.read()
