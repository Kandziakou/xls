import os.path
from pprint import pprint
from random import randint

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

path = os.path.join(os.curdir, "files")
name = 'products_selected.xlsx'

required_to_fill = ''

categories = ('bool_values', 'CMFR', 'RiskCategory', 'Country', 'BrandGroup', 'TransportConditions', 'ProductUnit',
              'ProductPackingType', 'FNS', 'TNVED', 'OKPD', 'VAT', 'packagepack_type', 'packagepack_material',)


def to_xlsx(file):
    pass


def to_xls(file):
    pass


class Xls:
    def __init__(self, path: str = '', name: str = '', required: list = None):
        self.__file = os.path.join(path, name)
        self.__sheet = None
        self.__cell = None
        self.__required = required
        self.__rows = []

    @property
    def sheet(self):
        self.__sheet = openpyxl.open(self.__file)['ФПИ']
        return self

    def cell(self, cell: str = 'A1'):
        self.__cell = self.__sheet[cell]
        return self.__cell

    def find(self, value, sheet: Worksheet, from_row: int = 1, col: int = 1) -> tuple[int, int]:
        for row in range(1, sheet.max_row):
            if sheet.cell(row, col).value == value:
                return row, col
        return sheet.max_row, col

    @property
    def color(self) -> str:
        return self.__cell.fill.fgColor.rgb

    @property
    def required(self):
        if self.__required is None:
            self.__required = ['A1']
        return self.__required

    def row(self, index: int = None):
        rows = self.__sheet.rows
        for i in range(index):
            self.__rows.append(next(rows))
        return [[i.value for i in self.__rows[j]] for j in range(len(self.__rows))]

    def product_category(self, value):
        """
        select 1lvl category C1 from :categories: sheet
        select 2lvl category C2 from :categories2: sheet (from list of pairs C1-C2 in sheet)
        select 3lvl category C3 from :categories3: sheet (from list of pairs C2-C3 in sheet) (try if exists)
        select 4lvl category C4 from :categories4: sheet (from list of pairs C3-C4 in sheet) (try if exists)
        :param value:
        :return:
        """

    def dropdown(self, category, next_category, col: int = 1, rowx: int = 1, value: str = None):
        sheet = openpyxl.open(self.__file)['catalogs']
        start = self.find(category, sheet)[0]
        end = self.find(next_category, sheet, start)[0]
        if category is None:
            rowx = randint(start, end)
        elif category is not None:
            rowx = self.find(value, sheet)[0]
            if rowx < start or rowx > end:
                rowx = randint(start, end)
        return sheet.cell(rowx, col)

    def input(self, value):
        pass


def test_open():
    file = Xls(path, name)
    sheet = file.sheet
    #return sheet.cell('E5').value
    #return sheet.row(5)
    #todo написать функции для ввода в нужную ячейку а также алгоритмы для базовы случаев- общий случай, молочка, алко и товары животного происхождения
    return file.dropdown(categories[4], categories[5], value='Продукция низкого риска').value


if __name__ == '__main__':
    pprint(test_open())
