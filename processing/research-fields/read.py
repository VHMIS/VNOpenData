# Chương trình đọc dữ liệu từ file excel
# File excel được biên tập lại từ quyết định 12/2008/QĐ-BKHCN của bộ khoa học công nghệ

# This script reads data from excel file that is edited from document 12/2008/QĐ-BKHCN by Ministry of Science and Technology

import io, json, xlrd

# open excel, get sheet
book = xlrd.open_workbook("research_fields_2008.xls")
sheet = book.sheet_by_index(0)

# init data dictionary

data = {
    'vhdata': {
        'research-fields': {
            'level1': {
            },
            'name-mapping': {
                'level1': 'Lĩnh vực',
                'level2': 'Ngành',
                'level3': 'Chuyên ngành'
            }
        }
    }
}

# Example data
# data['vhdata']['administrative-divisions']['province']['01'] = {
#     'name': 'Thành phố Hà Nội',
#     'code': '01',
#     'full-name': 'Thành phố Hà Nội',
#     'full-code': '01',
#     'district': {
#         '01': {
#             'name' : 'Quận Tây Hồ',
#             'full-name': 'Quận Tây Hồ - Thành phố Hà Nội',
#             'code': '01',
#             'full-code': '01.01',
#             'commune': {
#                 '001' : {
#                     'name' : 'Phường Nhân Chính',
#                     'full-name' : 'Phường Nhân Chính - Quận Tây Hồ - Thành phố Hà Nội',
#                     'code' : '001',
#                     'full-code' : '01.01.001',
#                 }
#             }
#         }
#     }
# }


# track code
currentLevel1Code = ''
currentLevel2Code = ''

# in xldr, row and col use 0-index, so cell with row 1 and col 1 means 'B2' cell
rowStart = 1 #row 2
rowEnd = sheet.nrows

# read data row by row
for row in range(rowStart, rowEnd):
    level1Code = str(sheet.cell_value(row, 0))
    level2Code = str(sheet.cell_value(row, 1))
    level3Code = str(sheet.cell_value(row, 2))
    name = sheet.cell_value(row, 3)
    note = sheet.cell_value(row, 4)

    if level3Code != '':
        level1Code = level3Code[0:1]
        level2Code = level3Code[1:3]
        fullCode = level3Code[0:5]
        level3Code = level3Code[3:5]
    elif level2Code != '':
        level1Code = level2Code[0:1]
        fullCode = level2Code[0:3]
        level2Code = level2Code[1:3]
    else:
        fullCode = level1Code[0:1]
        level1Code = level1Code[0:1]
    
    # Check if province code is not existed
    if currentLevel1Code != level1Code:
        currentLevel1Code = level1Code
        currentLevel2Code = ''
        data['vhdata']['research-fields']['level1'][currentLevel1Code] = {
            'name': name.capitalize(),
            'note': note,
            'code': level1Code,
            'full-code': fullCode,
            'level2': {}
        }
        continue
    
    # Check if district code is not exist
    if currentLevel2Code != level2Code:
        currentLevel2Code = level2Code
        data['vhdata']['research-fields']['level1'][currentLevel1Code]['level2'][currentLevel2Code] = {
            'name': name.capitalize(),
            'note': note,
            'code': level2Code,
            'full-code': fullCode,
            'level3': {}
        }
        continue
        
    data['vhdata']['research-fields']['level1'][currentLevel1Code]['level2'][currentLevel2Code]['level3'][level3Code] = {
        'name': name,
        'note': note,
        'full-name': note,
        'code': level3Code,
        'full-code': fullCode
    }

# export json file
with open('../../data/research-fields/full.json', 'w') as f:
    json.dump(data, f, indent=4, sort_keys=True)

# end
