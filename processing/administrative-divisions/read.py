# Chương trình đọc dữ liệu từ file excel
# File excel được lấy từ trang web của tổng cục thống kê

# This script reads data from excel file that is downloaded from website of GENERAL STATISTICS OFFICE of VIETNAM

import io, json, xlrd

# open excel, get sheet
book = xlrd.open_workbook("donvi_hanhchinh.xls")
sheet = book.sheet_by_index(0)

# init data dictionary

data = {
    'vhdata': {
        'administrative-divisions': {
            'province': {
                
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
currentProvinceCode = ''
currentDistrictCode = ''

# in xldr, row and col use 0-index, so cell with row 1 and col 1 means 'B2' cell
rowStart = 1 #row 2
rowEnd = sheet.nrows

# read data row by row
for row in range(rowStart, rowEnd):
    provinceCode = sheet.cell_value(row, 0)
    districtCode = sheet.cell_value(row, 2)
    communeCode = sheet.cell_value(row, 4)
    provinceName = sheet.cell_value(row, 1)
    districtName = sheet.cell_value(row, 3)
    communeName = sheet.cell_value(row, 5)
    
    # Check if province code is not existed
    if currentProvinceCode != provinceCode:
        currentProvinceCode = provinceCode
        currentDistrictCode = ''
        currentCommuneCode = ''
        data['vhdata']['administrative-divisions']['province'][currentProvinceCode] = {
            'name': provinceName,
            'full-name': provinceName,
            'code': provinceCode,
            'full-code': provinceCode,
            'district': {}
        }
    
    # Check if district code is not exist
    if currentDistrictCode != districtCode:
        currentDistrictCode = districtCode
        data['vhdata']['administrative-divisions']['province'][currentProvinceCode]['district'][currentDistrictCode] = {
            'name': districtName,
            'full-name': districtName + ' - ' + provinceName,
            'code': districtCode,
            'full-code': provinceCode + '.' + districtCode,
            'commune': {}
        }
        
    data['vhdata']['administrative-divisions']['province'][currentProvinceCode]['district'][currentDistrictCode]['commune'][communeCode] = {
        'name': communeName,
        'full-name': communeName + ' - ' + districtName + ' - ' + provinceName,
        'code': communeCode,
        'full-code': provinceCode + '.' + districtCode + '.' + communeCode
    }

# export json file
with open('../../data/administrative-divisions/full.json', 'w') as f:
    json.dump(data, f, indent=4, sort_keys=True)

# end