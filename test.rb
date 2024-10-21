require 'watir'
require 'win32ole'

# Khởi tạo browser
browser = Watir::Browser.new :edge

# Kết nối đến Excel
excel = WIN32OLE.new('Excel.Application')
workbook = excel.Workbooks.Open('C:\watir\WatirTestData.xlsx')
worksheet = workbook.Worksheets('Sheet1')

# Mở trang web
browser.goto 'http://127.0.0.1:5500/index.html'

# Xác định số hàng có dữ liệu trong Excel
row_count = worksheet.UsedRange.Rows.Count

# Duyệt qua từng hàng để chạy các trường hợp kiểm thử
(2..row_count).each do |row|
  # Đọc dữ liệu từ Excel
  email = worksheet.Cells(row, 'A').value
  password = worksheet.Cells(row, 'B').value

  # Điền dữ liệu vào form đăng nhập
  browser.text_field(id: 'email').set(email)
  browser.text_field(id: 'password').set(password)

  # Nhấn nút đăng nhập
  browser.button(xpath: '/html/body/div/button').click

  # Tạm dừng 2 giây để chờ kết quả xử lý
  sleep(2)

  # Lấy kết quả và ghi lại vào Excel
  worksheet.Cells(row, 'C').value = browser.div(id: 'emailResult').text
  worksheet.Cells(row, 'D').value = browser.div(id: 'passwordResult').text
  worksheet.Cells(row, 'E').value = browser.div(id: 'loginResult').text
end

# Lưu file Excel và đóng
workbook.Save
workbook.Close
excel.Quit

# Đóng browser
browser.close
