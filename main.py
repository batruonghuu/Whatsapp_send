
import pandas as pd
import pywhatkit
from tkinter import filedialog

# Lấy đường dẫn tệp tin từ người dùng
filename = filedialog.askopenfilename()

# Mở workbook và chọn sheet
xl = pd.ExcelFile(filename)
sheet_name = xl.sheet_names[0]

# Đọc sheet vào DataFrame
df = xl.parse(sheet_name)

# Lặp qua từng hàng trong DataFrame
for index, row in df.iterrows():
    phone_number = str(row[1])
    message = str(row[2])
    print(phone_number + " " + message)
    # Gửi tin nhắn qua WhatsApp
    pywhatkit.sendwhatmsg(phone_number, message, 0, 0)

# Đóng workbook
xl.close()