require 'watir'
require 'win32ole'

# Khởi tạo browser
browser = Watir::Browser.new

#Kết nối đến Excel
excel = WIN32OLE.connect('C:\watir\WatirTestData.xlsx')
worksheet = excel.Worksheets('Sheet1')

# Mở trang web linked in
browser.goto 'http://127.0.0.1:5500/index.html'


# trường hợp 1: tất cả các trường đăng nhập đều rỗng
# browser.text_field(id: 'username').set(worksheet.Cells(11, 'A').value)
# browser.text_field(id: 'password').set(worksheet.Cells(11, 'B').value)

# browser.button(xpath: '//*[@id="organic-div"]/form/div[3]/button').click

# # trả kết quả về excel
# worksheet.Cells(11, 'C').value = browser.div(id: 'error-for-username').text

# #Trường hợp 2: chỉ nhập username
# browser.text_field(id: 'username').set(worksheet.Cells(10, 'A').value)
# browser.text_field(id: 'password').set(worksheet.Cells(10, 'B').value)

# browser.button(xpath: '//*[@id="organic-div"]/form/div[3]/button').click
# # trả kết quả về excel
# worksheet.Cells(10, 'C').value = browser.div(id: 'error-for-password').text

#Trường hợp 3: passwork nhập sai định dạng
browser.text_field(id: 'email').set(worksheet.Cells(9, 'A').value)
browser.text_field(id: 'password').set(worksheet.Cells(9, 'B').value)

browser.button(xpath: '/html/body/div/button').click
# trả kết quả về excel
worksheet.Cells(9, 'C').value = browser.div(id: 'emailResult').text
worksheet.Cells(9, 'D').value = browser.div(id: 'passwordResult').text
worksheet.Cells(9, 'E').value = browser.div(id: 'loginResult').text


# Tạm dừng 5 giây để chờ trang web load xong
sleep(1000)

browser.close()