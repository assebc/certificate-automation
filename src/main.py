import openpyxl
from PIL import ImageFont, Image, ImageDraw

workbook = openpyxl.load_workbook('../input/dataset_students.xlsx')
sheet = workbook['Sheet1']
name_font = ImageFont.truetype('./fonts/tahomabd.ttf')
font = ImageFont.truetype('./fonts/tahoma.ttf')
image = Image.open('./images/certificate.png')
draw = ImageDraw.Draw(image)

for row in sheet.iter_rows(min_row=2):
  course_name = row[0].value
  student_name = row[1].value
  participation_type = row[2].value
  start_date = row[3].value
  end_date = row[4].value
  emit_date = row[6].value
  grade = row[7].value

  if grade > 10:
    draw.text((100, 100), student_name, font=name_font, fill='black')
    draw.text((100, 200), course_name, font=font, fill='black')
    draw.text((100, 300), str(grade), font=font, fill='black')
    draw.text((100, 400), emit_date, font=font, fill='black')
    image.save(f'../output/certificate_{course_name}/{student_name}.png')