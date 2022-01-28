import pywhatkit
import time
import openpyxl
from pathlib import Path

# xlsx_file = Path('Guests.xlsx')
xlsx_file = Path('Test.xlsx')
wb_obj = openpyxl.load_workbook(xlsx_file)
sheet = wb_obj.active
rows = list(sheet.iter_rows())[2:]

name_index = 0
phone_index = 4
with_out_point_index = 9
israel_area = "+972"
text_message = " היקרים!\nאנחנו שמחים ונרגשים להזמינכם לחגוג עימנו את יום נישואינו!"

names = []
phone_numbers = []

for row in rows:
    if row[name_index].value is not None and row[phone_index].value is not None:
        names.append(row[name_index].value)
        clean_phone_number = str(row[phone_index].value)[:with_out_point_index]
        with_are_code = israel_area + str(clean_phone_number)
        phone_numbers.append(with_are_code)

        print(str(row[name_index].value) + ": " + str(with_are_code))
print(len(phone_numbers))
guests = zip(names, phone_numbers)
for guest in guests:
    print(guest[0])
    pywhatkit.sendwhats_image(guest[1], "WeddingInvitation.jpeg", guest[0] + text_message, 20, True, 5)
    time.sleep(1)
