# 引入opnepyxl来解析excel
import openpyxl as xl
import os

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


# 在对应的cell中写入内容
def write_cell(excel_sheet, cell_position, content):
    excel_sheet[cell_position].value = content


# 更改文件之后必须调用
def update_excel(workbook, excel_name):
    os.remove(excel_name)
    workbook.save(excel_name)

# 测试的main函数
def test_code():
    excel_name = 'test.xlsx'
    workbook = get_excel(excel_name)
    sheet = get_sheet(workbook, 'Sheet2')
    cells = get_cells(sheet, ['a1', 'a5'])
    write_cell(sheet, 'a11', 'aaaaaaa')
    update_excel(workbook, excel_name)


test_code()
