import xlsxwriter
import openpyxl

xb = openpyxl.load_workbook('1-12.xlsx')

A1 = xb['！A1']

original_list = []

for i in range(1, 900):
    original_list.append(str(A1.cell(row=i, column=1).value).replace('\xa0', ' '))

# print(a)

first_handle_list = []
s = ''
for i in range(0, len(original_list)):
    if original_list[i] == 'None':
        first_handle_list.append(s)
        first_handle_list.append('None')
        s = ''
    elif original_list[i].find('2022') == -1:
        s += original_list[i] + ' '
    else:
        temp = original_list[i].split(' ')
        for j in range(0, len(temp)):
            first_handle_list.append(temp[j])
# print(first_handle_list)

second_handle_list = []
for i in range(0, len(first_handle_list)):
    if first_handle_list[i] != '' and first_handle_list[i].find('撤回了一条消息') == -1 and \
            first_handle_list[i].find('修改了群名称') == -1:
        second_handle_list.append(first_handle_list[i])
# print(second_handle_list)

final_handle_list = ['']
for i in range(0, len(second_handle_list)):
    if final_handle_list[len(final_handle_list) - 1] != second_handle_list[i]:
        final_handle_list.append(second_handle_list[i])
final_handle_list.remove('')
print(final_handle_list)


# 创建工作簿
file_name = "demo.xlsx"
workbook = xlsxwriter.Workbook(file_name)
# 创建工作表
worksheet = workbook.add_worksheet('first_sheet1')

line = []
row = 0
col = 0
i = 0
while i < len(final_handle_list):
    if final_handle_list[i] != 'None':
        line.append(final_handle_list[i])
        # print(line)
        i += 1
    else:
        worksheet.write_row(row, col, line)
        row += 1
        i += 1
        line.clear()

# 关闭工作簿
workbook.close()
