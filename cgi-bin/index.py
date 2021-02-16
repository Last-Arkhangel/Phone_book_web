#!/usr/bin/env python
# -*- coding: utf-8 -*-
import cgi
import html
import xlrd
import cssutils
import cgitb; cgitb.enable()

# Задаем переменные
form = cgi.FieldStorage()
a = form.getfirst("TEXT_1", "не задано")
b = form.getfirst("TEXT_2", "не задано")
a = html.escape(a)
b = html.escape(b)
a = a.title()
b = b.title()

print("Content-Type: text/html; charset=cp1251\r")
print('\r')
print("""<!DOCTYPE HTML>
<html>

<head>
<meta charset='UTF-8'>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Телефонний довідник</title>
<meta name="keywords" content='Телефонний довідник'>
<meta name="description" content='Телефонний довідник'>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="author" content="Denis" />
<link rel="shortcut icon" href="../img/phone.png"> 
<link rel="stylesheet" href="../css/main.css" type="text/css" charset="utf-8">    
</head>
<body>
<header>
    <div class="hed" id="page-wrad">
        <div class="hed-r">
            <img src="../img/phone.png" alt="Телефонний довідник">
        </div>	
        <div class="hed-l">
            <font>Телефонний довідник та дати народження</font>
        </div>
    </div>
</header>
<br>
<div class="form" id="page-wrap">
    <div>
        <form action="../cgi-bin/index.py" accept-charset="utf-8">
	        <input class="in-for" type="text" placeholder="Дані співробітника (1)..." name="TEXT_1">
	        <input class="sech-print" type="submit" value=" Пошук ">
	        <output type="text" name="OUT">
	    </form>
    </div>
</div>
<div class="form" id="page-wrad">
    <div>
        <form action="../cgi-bin/index.py" accept-charset="utf-8">
	        <input class="in-for" type="text" placeholder="Дати народження (2)..." name="TEXT_2">
	        <input class="sech-print" type="submit" value=" Пошук ">
	        <output type="text" name="OUT">
	    </form>
    </div>
</div>

<div id="page-wrad"><h7>Для пошуку ведіть (1) прізвище або ім'я, (2) прізвище або ім'я або місяць нароження. Мінімальна кількість символів для пошуку три.</h7></div>

<br>
<div id="page-wrad">
 """)
 
print("<h2>Дані співробітника (телефони)</h2>")

print("""
<table class="table table-borderedt table-bordc table-hovert">
      <thead>
        <tr>
          <th scope="col">Прізвище Ім'я По батькові</th>
          <th scope="col">Підрозділ</th>          
          <th scope="col">Посада</th>
          <th scope="col">Каб.</th>
          <th scope="col">Тел. внутр.</th>
          <th scope="col">Тел. місцев.</th>
          <th scope="col">Тел. моб. 1</th>
          <th scope="col">Тел. моб. 2</th>
        </tr>
      </thead>
      <tbody>
 """)     

# Код телефоний довідник
rb = xlrd.open_workbook(r'./Телефонний довідник.xlsx', encoding_override='utf-8')
sheet = rb.sheet_by_index(0)
vals = [sheet.row_values(rownum) for rownum in range(1, sheet.nrows)]
for otdel, kabinet, posada, last_name, first_name, third_name, data_r, phone_v, phone_g, phone_m1, phone_m2, data_p, data_z, mesac_r in vals:
    first_name = first_name.strip()
    if a == last_name[:3] or a == last_name[:4] or a == last_name[:5] \
            or a == last_name[:6] or a == last_name[:7] or a == last_name[:8] \
            or a == last_name[:9] or a == last_name[:10] or a == last_name[:11] \
            or a == last_name[:12] or a == last_name[:13] or a == last_name[:14] \
            or a == last_name[:15] or a == last_name[:17] or a == last_name[:17] \
            or a == first_name[:3] or a == first_name[:4] or a == first_name[:5] \
            or a == first_name[:6] or a == first_name[:7] or a == first_name[:8] \
            or a == first_name[:9] or a == first_name[:10] or a == first_name[:11] \
            or a == first_name[:12] or a == first_name[:13] or a == first_name[:14] \
            or a == first_name[:15] or a == first_name[:16] or a == first_name[:17]:
        kabinet = str(kabinet)
        kabinet = kabinet[:4]
        phone_v = str(phone_v)
        phone_v = phone_v[:3]
        phone_g = str(phone_g)
        phone_g = phone_g[:21]
        phone_m1 = str(phone_m1)
        phone_m1 = phone_m1[:21]
        phone_m2 = str(phone_m2)
        phone_m2 = phone_m2[:21]
        data_r = str(data_r)
        data_r = data_r[:21]
        data_p = str(data_p)
        data_p = data_p[:21]
        data_z = str(data_z)
        data_z = data_z[:21]
        print('<tr><td>{} {} {}</td> <td>{}</td> <td>{}</td> <td>{}</td> <td>{}</td> <td>{}</td> <td>{}</td> <td>{}</td></tr>'.format(last_name, first_name, third_name, otdel, posada, kabinet, phone_v, phone_g, phone_m1, phone_m2) + ' ')
        
    elif a != last_name:
        pass

print("""
      </tbody>
  </table>  
""")

print("<br><h2>Дати народження, прийняття, звільнення</h2>")

print("""
<table class="table table-borderedt table-bordd table-hovert">
      <thead>
        <tr>
          <th scope="col">Прізвище Ім'я По батькові</th>
          <th scope="col">Посада</th>
          <th scope="col">Дата народження</th>
          <th scope="col">З якого часу працює</th>
          <th scope="col">До якого часу працював</th>
        </tr>
      </thead>
      <tbody>
 """)  
        
# Код дата народження
rb = xlrd.open_workbook(r'./Телефонний довідник.xlsx', encoding_override='utf-8')
sheet = rb.sheet_by_index(0)
vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
for otdel, kabinet, posada, last_name, first_name, third_name, data_r, phone_v, phone_g, phone_m1, phone_m2, data_p, data_z, mesac_r in vals:
    first_name = first_name.strip()
    if b == last_name[:3] or b == last_name[:4] or b == last_name[:5] \
            or b == last_name[:6] or b == last_name[:7] or b == last_name[:8] \
            or b == last_name[:9] or b == last_name[:10] or b == last_name[:11] \
            or b == last_name[:12] or b == last_name[:13] or b == last_name[:14] \
            or b == last_name[:15] or b == last_name[:16] or b == last_name[:17] \
            or b == first_name[:3] or b == first_name[:4] or b == first_name[:5] \
            or b == first_name[:6] or b == first_name[:7] or b == first_name[:8] \
            or b == first_name[:9] or b == first_name[:10] or b == first_name[:11] \
            or b == first_name[:12] or b == first_name[:13] or b == first_name[:14] \
            or b == first_name[:15] or b == first_name[:16] or b == first_name[:17] \
            or b == mesac_r[:3] or b == mesac_r[:4] or b == mesac_r[:5] \
            or b == mesac_r[:6] or b == mesac_r[:7] or b == mesac_r[:8] \
            or b == mesac_r[:9] or b == mesac_r[:10] or b == mesac_r[:11]:
        kabinet = str(kabinet)
        kabinet = kabinet[:4]
        phone_v = str(phone_v)
        phone_v = phone_v[:3]
        phone_g = str(phone_g)
        phone_g = phone_g[:21]
        phone_m1 = str(phone_m1)
        phone_m1 = phone_m1[:21]
        phone_m2 = str(phone_m2)
        phone_m2 = phone_m2[:21]
        data_r = str(data_r)
        data_r = data_r[:21]
        data_p = str(data_p)
        data_p = data_p[:21]
        data_z = str(data_z)
        data_z = data_z[:21]
        print('<tr><td><span style="color: #cc0000;"><b>{} {} {}</b></span></td> <td>{}</td> <td><span style="color: #cc0000;"><b>{}</b></span></td> <td>{}</td> <td>{}</td> </tr>'.format(last_name, first_name, third_name, posada, data_r, data_p, data_z) + ' ')
        
    elif b != first_name:
        pass

print("""
      </tbody>
  </table>  
""")

print("<br><br><p style='text-align: center;'><span style='font-size: 10pt;'>Авторське право Denis © 2020. Всі права захищені.</span></p>")

# Закрываем html
print("""</div>
</body>
</html>""")