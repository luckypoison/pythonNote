# 引入opnepyxl来解析excel
import openpyxl as xl


# 通过excel名字获取excel的workbook
def get_excel(excel_name):
    return xl.load_workbook(excel_name)


# 通过excel的workbook，和sheet的名称获取sheet
def get_sheet(excel_workbook, sheet_name):
    return excel_workbook[sheet_name]


# 通过excel的sheet，和cell的位置获取cell（如果为第一行第一列，位置为a1）
def get_cells(excel_sheet, cell_positions):
    cells = []
    for position in cell_positions:
        cells.append(excel_sheet[position])
    return cells


# 测试的main函数
def test_code():
    workbook = get_excel('test.xlsx')
    sheet = get_sheet(workbook, 'Sheet2')
    cells = get_cells(sheet, ['a1', 'a5'])
    print(cells[0].value)
    print(cells[1].value)


test_code()
