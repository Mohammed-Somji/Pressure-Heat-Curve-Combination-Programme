# Pressure-Heat-Curve-Combination-Programme

#In chemical engineering, fluids are often represented in what's known as heat curves where information about its properties is shown. This is a programme that allows you to combine heat curve properties at two different pressure levels. 

#The reason for making this programme was due to a company I used to work for. Where their pogramme could only use one heat curve per fluid. However their customers would often send in heat curves at multiple pressures. 

#Therefore this programme provided a quick and easy way to combine heat curves in a matter of seconds, rather than hours or days via linear interpolations of data.

#This programme was written in python, and is conncected to excel via the import openpyxl command. 

import openpyxl

wb = openpyxl.load_workbook('ZIP.xlsx')

sheet = wb['Sheet']


y = sheet['B4'].value


sheet = wb['Sheet1']
for i in range(29, 40):
    if i == 29:
        i = 'AC'
    elif i == 30:
        i = 'AD'
    elif i == 31:
        i = 'AE'
    elif i == 32:
        i = 'Af'
    elif i == 33:
        i = 'AG'
    elif i == 34:
        i = 'AH'
    elif i == 35:
        i = 'AI'
    elif i == 36:
        i = 'AJ'
    elif i == 37:
        i = 'AK'
    elif i == 38:
        i = 'AL'
    elif i == 39:
        i = 'AM'
    else:
        ''
    for q in range(8, 8 + y):
        sheet[str(i) + str(q)].value = ' '


sheet = wb['Sheet']


# Code for presenting temperature data as per range from the customer (table 1, temperatures)
#-------------------------------------------------------------------------------------------------------------------
for o in range(8, (8 + y)):
    a = sheet.cell(row=o, column=2).value
    b = sheet.cell(row=o + 1, column=2).value
    c = sheet.cell(row=o - 1, column=2).value
    if sheet['B2'].value >= a >= sheet['B3'].value:
        if sheet['B' + str(o)].value < sheet['B2'].value < sheet['B' + str(o - 1)].value:
            sheet['N' + str(o)].value = a
        else:
            if sheet['B' + str(o)].value < sheet['B3'].value < sheet['B' + str(o - 1)].value:
                sheet['N' + str(o + 1)].value = sheet['B3'].value
            else:
                sheet['N' + str(o)].value = a
    else:
        if a > sheet['B2'].value > b:
            sheet['N' + str(o)].value = sheet['B2'].value
        else:
            if a < sheet['B3'].value < c:
                sheet['N' + str(o)].value = sheet['B3'].value
            else:
                ''
#-------------------------------------------------------------------------------------------------------------------


# Code for interpolating data as per range from the customer (Table 1, Enthalpy)
#-------------------------------------------------------------------------------------------------------------------
for x in range(8, (7 + y)):
    a = sheet.cell(row=x, column=2).value
    b = sheet.cell(row=x + 1, column=2).value
    c = sheet.cell(row=x, column=14).value
    d = sheet.cell(row=x, column=3).value
    e = sheet.cell(row=x + 1, column=3).value
    if isinstance(sheet['N' + str(x)].value, (int, float)):
        if a >= sheet['N' + str(x)].value >= b:
            sheet['O' + str(x)].value = ((d - e) / (a - b)) * c + (d - ((d - e) / (a - b)) * a)
        else:
            ''
    else:
        ''
for x1 in range(9, (8 + y)):
    a = sheet.cell(row=x1, column=2).value
    b = sheet.cell(row=x1 - 1, column=2).value
    c = sheet.cell(row=x1, column=14).value
    d = sheet.cell(row=x1, column=3).value
    e = sheet.cell(row=x1 - 1, column=3).value
    if isinstance(sheet['N' + str(x1 - 1)].value, (int, float)):
        if b >= sheet['N' + str(x1 - 1)].value > a:
            sheet['O' + str(x1)].value = ((e - d) / (b - a)) * c + (e - ((e - d) / (b - a)) * b)
        else:
            ''
    else:
        ''
#-------------------------------------------------------------------------------------------------------------------


# Code for interpolating data as per range from the customer (Table 1, Vapour Fraction)
#-------------------------------------------------------------------------------------------------------------------
for x in range(8, (7 + y)):
    a = sheet.cell(row=x, column=2).value
    b = sheet.cell(row=x + 1, column=2).value
    c = sheet.cell(row=x, column=14).value
    d = sheet.cell(row=x, column=4).value
    e = sheet.cell(row=x + 1, column=4).value
    if isinstance(sheet['N' + str(x)].value, (int, float)):
        if a >= sheet['N' + str(x)].value >= b:
            sheet['P' + str(x)].value = ((d - e) / (a - b)) * c + (d - ((d - e) / (a - b)) * a)
        else:
            ''
    else:
        ''
for x1 in range(9, (8 + y)):
    a = sheet.cell(row=x1, column=2).value
    b = sheet.cell(row=x1 - 1, column=2).value
    c = sheet.cell(row=x1, column=14).value
    d = sheet.cell(row=x1, column=4).value
    e = sheet.cell(row=x1 - 1, column=4).value
    if isinstance(sheet['N' + str(x1 - 1)].value, (int, float)):
        if b >= sheet['N' + str(x1 - 1)].value > a:
            sheet['P' + str(x1)].value = ((e - d) / (b - a)) * c + (e - ((e - d) / (b - a)) * b)
        else:
            ''
    else:
        ''


# Code for interpolating data as per range from the customer (Table 1, Vapour, Density)
#-------------------------------------------------------------------------------------------------------------------
for x in range(8, (7 + y)):
    a = sheet.cell(row=x, column=2).value
    b = sheet.cell(row=x + 1, column=2).value
    c = sheet.cell(row=x, column=14).value
    d = sheet.cell(row=x, column=5).value
    e = sheet.cell(row=x + 1, column=5).value
    if isinstance(sheet['N' + str(x)].value, (int, float)):
        if a >= sheet['N' + str(x)].value >= b:
            sheet['Q' + str(x)].value = ((d - e) / (a - b)) * c + (d - ((d - e) / (a - b)) * a)
        else:
            ''
    else:
        ''
for x1 in range(9, (8 + y)):
    a = sheet.cell(row=x1, column=2).value
    b = sheet.cell(row=x1 - 1, column=2).value
    c = sheet.cell(row=x1, column=14).value
    d = sheet.cell(row=x1, column=5).value
    e = sheet.cell(row=x1 - 1, column=5).value
    if isinstance(sheet['N' + str(x1 - 1)].value, (int, float)):
        if b >= sheet['N' + str(x1 - 1)].value > a:
            sheet['Q' + str(x1)].value = ((e - d) / (b - a)) * c + (e - ((e - d) / (b - a)) * b)
        else:
            ''
    else:
        ''
#-------------------------------------------------------------------------------------------------------------------


# Code for interpolating data as per range from the customer (Table 1, Vapour, Heat Capacity)
#-------------------------------------------------------------------------------------------------------------------
for x in range(8, (7 + y)):
    a = sheet.cell(row=x, column=2).value
    b = sheet.cell(row=x + 1, column=2).value
    c = sheet.cell(row=x, column=14).value
    d = sheet.cell(row=x, column=6).value
    e = sheet.cell(row=x + 1, column=6).value
    if isinstance(sheet['N' + str(x)].value, (int, float)):
        if a >= sheet['N' + str(x)].value >= b:
            sheet['R' + str(x)].value = ((d - e) / (a - b)) * c + (d - ((d - e) / (a - b)) * a)
        else:
            ''
    else:
        ''
for x1 in range(9, (8 + y)):
    a = sheet.cell(row=x1, column=2).value
    b = sheet.cell(row=x1 - 1, column=2).value
    c = sheet.cell(row=x1, column=14).value
    d = sheet.cell(row=x1, column=6).value
    e = sheet.cell(row=x1 - 1, column=6).value
    if isinstance(sheet['N' + str(x1 - 1)].value, (int, float)):
        if b >= sheet['N' + str(x1 - 1)].value > a:
            sheet['R' + str(x1)].value = ((e - d) / (b - a)) * c + (e - ((e - d) / (b - a)) * b)
        else:
            ''
    else:
        ''
#-------------------------------------------------------------------------------------------------------------------


# Code for interpolating data as per range from the customer (Table 1, Vapour, Viscocity)
#-------------------------------------------------------------------------------------------------------------------
for x in range(8, (7 + y)):
    a = sheet.cell(row=x, column=2).value
    b = sheet.cell(row=x + 1, column=2).value
    c = sheet.cell(row=x, column=14).value
    d = sheet.cell(row=x, column=7).value
    e = sheet.cell(row=x + 1, column=7).value
    if isinstance(sheet['N' + str(x)].value, (int, float)):
        if a >= sheet['N' + str(x)].value >= b:
            sheet['S' + str(x)].value = ((d - e) / (a - b)) * c + (d - ((d - e) / (a - b)) * a)
        else:
            ''
    else:
        ''
for x1 in range(9, (8 + y)):
    a = sheet.cell(row=x1, column=2).value
    b = sheet.cell(row=x1 - 1, column=2).value
    c = sheet.cell(row=x1, column=14).value
    d = sheet.cell(row=x1, column=7).value
    e = sheet.cell(row=x1 - 1, column=7).value
    if isinstance(sheet['N' + str(x1 - 1)].value, (int, float)):
        if b >= sheet['N' + str(x1 - 1)].value > a:
            sheet['S' + str(x1)].value = ((e - d) / (b - a)) * c + (e - ((e - d) / (b - a)) * b)
        else:
            ''
    else:
        ''
#-------------------------------------------------------------------------------------------------------------------


# Code for interpolating data as per range from the customer (Table 1, Vapour, Thermal Conductivity)
#-------------------------------------------------------------------------------------------------------------------
for x in range(8, (7 + y)):
    a = sheet.cell(row=x, column=2).value
    b = sheet.cell(row=x + 1, column=2).value
    c = sheet.cell(row=x, column=14).value
    d = sheet.cell(row=x, column=8).value
    e = sheet.cell(row=x + 1, column=8).value
    if isinstance(sheet['N' + str(x)].value, (int, float)):
        if a >= sheet['N' + str(x)].value >= b:
            sheet['T' + str(x)].value = ((d - e) / (a - b)) * c + (d - ((d - e) / (a - b)) * a)
        else:
            ''
    else:
        ''
for x1 in range(9, (8 + y)):
    a = sheet.cell(row=x1, column=2).value
    b = sheet.cell(row=x1 - 1, column=2).value
    c = sheet.cell(row=x1, column=14).value
    d = sheet.cell(row=x1, column=8).value
    e = sheet.cell(row=x1 - 1, column=8).value
    if isinstance(sheet['N' + str(x1 - 1)].value, (int, float)):
        if b >= sheet['N' + str(x1 - 1)].value > a:
            sheet['T' + str(x1)].value = ((e - d) / (b - a)) * c + (e - ((e - d) / (b - a)) * b)
        else:
            ''
    else:
        ''
#-------------------------------------------------------------------------------------------------------------------


# Code for interpolating data as per range from the customer (Table 1, Liquid, Density)
#-------------------------------------------------------------------------------------------------------------------
for x in range(8, (7 + y)):
    a = sheet.cell(row=x, column=2).value
    b = sheet.cell(row=x + 1, column=2).value
    c = sheet.cell(row=x, column=14).value
    d = sheet.cell(row=x, column=9).value
    e = sheet.cell(row=x + 1, column=9).value
    if isinstance(sheet['N' + str(x)].value, (int, float)):
        if a >= sheet['N' + str(x)].value >= b:
            sheet['U' + str(x)].value = ((d - e) / (a - b)) * c + (d - ((d - e) / (a - b)) * a)
        else:
            ''
    else:
        ''
for x1 in range(9, (8 + y)):
    a = sheet.cell(row=x1, column=2).value
    b = sheet.cell(row=x1 - 1, column=2).value
    c = sheet.cell(row=x1, column=14).value
    d = sheet.cell(row=x1, column=9).value
    e = sheet.cell(row=x1 - 1, column=9).value
    if isinstance(sheet['N' + str(x1 - 1)].value, (int, float)):
        if b >= sheet['N' + str(x1 - 1)].value > a:
            sheet['U' + str(x1)].value = ((e - d) / (b - a)) * c + (e - ((e - d) / (b - a)) * b)
        else:
            ''
    else:
        ''
#-------------------------------------------------------------------------------------------------------------------


# Code for interpolating data as per range from the customer (Table 1, Liquid, Heat Capacity)
#-------------------------------------------------------------------------------------------------------------------
for x in range(8, (7 + y)):
    a = sheet.cell(row=x, column=2).value
    b = sheet.cell(row=x + 1, column=2).value
    c = sheet.cell(row=x, column=14).value
    d = sheet.cell(row=x, column=10).value
    e = sheet.cell(row=x + 1, column=10).value
    if isinstance(sheet['N' + str(x)].value, (int, float)):
        if a >= sheet['N' + str(x)].value >= b:
            sheet['V' + str(x)].value = ((d - e) / (a - b)) * c + (d - ((d - e) / (a - b)) * a)
        else:
            ''
    else:
        ''
for x1 in range(9, (8 + y)):
    a = sheet.cell(row=x1, column=2).value
    b = sheet.cell(row=x1 - 1, column=2).value
    c = sheet.cell(row=x1, column=14).value
    d = sheet.cell(row=x1, column=10).value
    e = sheet.cell(row=x1 - 1, column=10).value
    if isinstance(sheet['N' + str(x1 - 1)].value, (int, float)):
        if b >= sheet['N' + str(x1 - 1)].value > a:
            sheet['V' + str(x1)].value = ((e - d) / (b - a)) * c + (e - ((e - d) / (b - a)) * b)
        else:
            ''
    else:
        ''
#-------------------------------------------------------------------------------------------------------------------


# Code for interpolating data as per range from the customer (Table 1, Liquid, Viscocity)
#-------------------------------------------------------------------------------------------------------------------
for x in range(8, (7 + y)):
    a = sheet.cell(row=x, column=2).value
    b = sheet.cell(row=x + 1, column=2).value
    c = sheet.cell(row=x, column=14).value
    d = sheet.cell(row=x, column=11).value
    e = sheet.cell(row=x + 1, column=11).value
    if isinstance(sheet['N' + str(x)].value, (int, float)):
        if a >= sheet['N' + str(x)].value >= b:
            sheet['W' + str(x)].value = ((d - e) / (a - b)) * c + (d - ((d - e) / (a - b)) * a)
        else:
            ''
    else:
        ''
for x1 in range(9, (8 + y)):
    a = sheet.cell(row=x1, column=2).value
    b = sheet.cell(row=x1 - 1, column=2).value
    c = sheet.cell(row=x1, column=14).value
    d = sheet.cell(row=x1, column=11).value
    e = sheet.cell(row=x1 - 1, column=11).value
    if isinstance(sheet['N' + str(x1 - 1)].value, (int, float)):
        if b >= sheet['N' + str(x1 - 1)].value > a:
            sheet['W' + str(x1)].value = ((e - d) / (b - a)) * c + (e - ((e - d) / (b - a)) * b)
        else:
            ''
    else:
        ''
#-------------------------------------------------------------------------------------------------------------------


# Code for interpolating data as per range from the customer (Table 1, Liquid, Thermal Conductivity)
#-------------------------------------------------------------------------------------------------------------------
for x in range(8, (7 + y)):
    a = sheet.cell(row=x, column=2).value
    b = sheet.cell(row=x + 1, column=2).value
    c = sheet.cell(row=x, column=14).value
    d = sheet.cell(row=x, column=12).value
    e = sheet.cell(row=x + 1, column=12).value
    if isinstance(sheet['N' + str(x)].value, (int, float)):
        if a >= sheet['N' + str(x)].value >= b:
            sheet['X' + str(x)].value = ((d - e) / (a - b)) * c + (d - ((d - e) / (a - b)) * a)
        else:
            ''
    else:
        ''
for x1 in range(9, (8 + y)):
    a = sheet.cell(row=x1, column=2).value
    b = sheet.cell(row=x1 - 1, column=2).value
    c = sheet.cell(row=x1, column=14).value
    d = sheet.cell(row=x1, column=12).value
    e = sheet.cell(row=x1 - 1, column=12).value
    if isinstance(sheet['N' + str(x1 - 1)].value, (int, float)):
        if b >= sheet['N' + str(x1 - 1)].value > a:
            sheet['X' + str(x1)].value = ((e - d) / (b - a)) * c + (e - ((e - d) / (b - a)) * b)
        else:
            ''
    else:
        ''
#-------------------------------------------------------------------------------------------------------------------


# Code for replicating Table 1 temperature steps to Table 2 (Table 2, Temperatures)
#-------------------------------------------------------------------------------------------------------------------
for x in range(55, (55 + y)):
    a = sheet.cell(row=x - 47, column=14).value
    sheet['N' + str(x)].value = a
#-------------------------------------------------------------------------------------------------------------------


# Code for interpolated values as per inlet pressure temperature steps (Table 2, Enthalpy)
#-------------------------------------------------------------------------------------------------------------------
for x in range (55, (55 + y)):
    a = sheet.cell(row = x, column = 14).value
    for z in range (55, (54 + y)):
        b = sheet.cell(row = z, column = 2).value
        c = sheet.cell(row = z + 1, column = 2).value
        d = sheet.cell(row = z, column = 3).value
        e = sheet.cell(row = z + 1, column = 3).value
        if b >= a >= c:
            sheet['O' + str(x)].value = ((d - e)/(b - c))*a + (d - ((d - e)/(b - c))*b)
        else:
            ''
#-------------------------------------------------------------------------------------------------------------------




# Code for interpolated values as per inlet pressure temperature steps (Table 2, Vapour Fraction)
#-------------------------------------------------------------------------------------------------------------------
for x in range (55, (55 + y)):
    a = sheet.cell(row = x, column = 14).value
    for z in range (55, (54 + y)):
        b = sheet.cell(row = z, column = 2).value
        c = sheet.cell(row = z + 1, column = 2).value
        d = sheet.cell(row = z, column = 4).value
        e = sheet.cell(row = z + 1, column = 4).value
        if b >= a >= c:
            sheet['P' + str(x)].value = ((d - e)/(b - c))*a + (d - ((d - e)/(b - c))*b)
        else:
            ''
#-------------------------------------------------------------------------------------------------------------------




# Code for interpolated values as per inlet pressure temperature steps (Table 2, Vapour, Density)
#-------------------------------------------------------------------------------------------------------------------
for x in range (55, (55 + y)):
    a = sheet.cell(row = x, column = 14).value
    for z in range (55, (54 + y)):
        b = sheet.cell(row = z, column = 2).value
        c = sheet.cell(row = z + 1, column = 2).value
        d = sheet.cell(row = z, column = 5).value
        e = sheet.cell(row = z + 1, column = 5).value
        if b >= a >= c:
            sheet['Q' + str(x)].value = ((d - e)/(b - c))*a + (d - ((d - e)/(b - c))*b)
        else:
            ''
#-------------------------------------------------------------------------------------------------------------------



# Code for interpolated values as per inlet pressure temperature steps (Table 2, Vapour, Heat Capacity)
#-------------------------------------------------------------------------------------------------------------------
for x in range (55, (55 + y)):
    a = sheet.cell(row = x, column = 14).value
    for z in range (55, (54 + y)):
        b = sheet.cell(row = z, column = 2).value
        c = sheet.cell(row = z + 1, column = 2).value
        d = sheet.cell(row = z, column = 6).value
        e = sheet.cell(row = z + 1, column = 6).value
        if b >= a >= c:
            sheet['R' + str(x)].value = ((d - e)/(b - c))*a + (d - ((d - e)/(b - c))*b)
        else:
            ''
#-------------------------------------------------------------------------------------------------------------------



# Code for interpolated values as per inlet pressure temperature steps (Table 2, Vapour, Viscocity)
#-------------------------------------------------------------------------------------------------------------------
for x in range (55, (55 + y)):
    a = sheet.cell(row = x, column = 14).value
    for z in range (55, (54 + y)):
        b = sheet.cell(row = z, column = 2).value
        c = sheet.cell(row = z + 1, column = 2).value
        d = sheet.cell(row = z, column = 7).value
        e = sheet.cell(row = z + 1, column = 7).value
        if b >= a >= c:
            sheet['S' + str(x)].value = ((d - e)/(b - c))*a + (d - ((d - e)/(b - c))*b)
        else:
            ''
#-------------------------------------------------------------------------------------------------------------------



# Code for interpolated values as per inlet pressure temperature steps (Table 2, Vapour, Thermal Conductivity)
#-------------------------------------------------------------------------------------------------------------------
for x in range (55, (55 + y)):
    a = sheet.cell(row = x, column = 14).value
    for z in range (55, (54 + y)):
        b = sheet.cell(row = z, column = 2).value
        c = sheet.cell(row = z + 1, column = 2).value
        d = sheet.cell(row = z, column = 8).value
        e = sheet.cell(row = z + 1, column = 8).value
        if b >= a >= c:
            sheet['T' + str(x)].value = ((d - e)/(b - c))*a + (d - ((d - e)/(b - c))*b)
        else:
            ''
#-------------------------------------------------------------------------------------------------------------------



# Code for interpolated values as per inlet pressure temperature steps (Table 2, Liquid, Density)
#-------------------------------------------------------------------------------------------------------------------
for x in range (55, (55 + y)):
    a = sheet.cell(row = x, column = 14).value
    for z in range (55, (54 + y)):
        b = sheet.cell(row = z, column = 2).value
        c = sheet.cell(row = z + 1, column = 2).value
        d = sheet.cell(row = z, column = 9).value
        e = sheet.cell(row = z + 1, column = 9).value
        if b >= a >= c:
            sheet['U' + str(x)].value = ((d - e)/(b - c))*a + (d - ((d - e)/(b - c))*b)
        else:
            ''
#-------------------------------------------------------------------------------------------------------------------



# Code for interpolated values as per inlet pressure temperature steps (Table 2, Liquid, Heat Capacity)
#-------------------------------------------------------------------------------------------------------------------
for x in range (55, (55 + y)):
    a = sheet.cell(row = x, column = 14).value
    for z in range (55, (54 + y)):
        b = sheet.cell(row = z, column = 2).value
        c = sheet.cell(row = z + 1, column = 2).value
        d = sheet.cell(row = z, column = 10).value
        e = sheet.cell(row = z + 1, column = 10).value
        if b >= a >= c:
            sheet['V' + str(x)].value = ((d - e)/(b - c))*a + (d - ((d - e)/(b - c))*b)
        else:
            ''
#-------------------------------------------------------------------------------------------------------------------



# Code for interpolated values as per inlet pressure temperature steps (Table 2, Liquid, Viscocity)
#-------------------------------------------------------------------------------------------------------------------
for x in range (55, (55 + y)):
    a = sheet.cell(row = x, column = 14).value
    for z in range (55, (54 + y)):
        b = sheet.cell(row = z, column = 2).value
        c = sheet.cell(row = z + 1, column = 2).value
        d = sheet.cell(row = z, column = 11).value
        e = sheet.cell(row = z + 1, column = 11).value
        if b >= a >= c:
            sheet['W' + str(x)].value = ((d - e)/(b - c))*a + (d - ((d - e)/(b - c))*b)
        else:
            ''
#-------------------------------------------------------------------------------------------------------------------




# Code for interpolated values as per inlet pressure temperature steps (Table 2, Liquid, Thermal Conductivity)
#-------------------------------------------------------------------------------------------------------------------
for x in range (55, (55 + y)):
    a = sheet.cell(row = x, column = 14).value
    for z in range (55, (54 + y)):
        b = sheet.cell(row = z, column = 2).value
        c = sheet.cell(row = z + 1, column = 2).value
        d = sheet.cell(row = z, column = 12).value
        e = sheet.cell(row = z + 1, column = 12).value
        if b >= a >= c:
            sheet['X' + str(x)].value = ((d - e)/(b - c))*a + (d - ((d - e)/(b - c))*b)
        else:
            ''
#-------------------------------------------------------------------------------------------------------------------


# Code for input of combined temperature steps (Table 3, Temperature)
#-------------------------------------------------------------------------------------------------------------------
for x in range (8, (8 + y)):
    a = sheet.cell(row = x, column = 14).value
    sheet['Z' + str(x)].value = a
#-------------------------------------------------------------------------------------------------------------------




# Code for input of combined heat curve (Table 3, Enthalpy)
#-------------------------------------------------------------------------------------------------------------------
if sheet['B1'].value == 'Duty':
    for x in range (0, y):
        a = (100 - x*((100)/(y - 1)))/100
        b = sheet.cell(row = x + 8, column = 15).value
        c = sheet.cell(row = x + 55, column = 15).value
        sheet['AA' + str(x + 8)].value = a*b + (1 - a)*c
elif sheet['B1'].value == 'Enthalpy':
    for x in range (0, y):
        a = (100 - x*((100)/(y - 1)))/100
        b = sheet.cell(row = x + 8, column = 15).value
        c = sheet.cell(row = x + 55, column = 15).value
        d = sheet['B5'].value
        e = sheet.cell(row = y + 54, column = 15).value
        sheet['AA' + str(x + 8)].value = (a*b + (1 - a)*c)*d - (e*d)
else:
    ''
#-------------------------------------------------------------------------------------------------------------------




#Code for input of combined heat curve (Table 3, Vapour Fraction)
#-------------------------------------------------------------------------------------------------------------------
for x in range (0, y):
    a = sheet.cell(row = x + 8, column = 16).value
    b = sheet.cell(row = x + 55, column = 16).value
    if x < (y-1)/2:
        sheet['AB' + str(x + 8)].value = a
    else:
        sheet['AB' + str(x + 8)].value = b
#-------------------------------------------------------------------------------------------------------------------




#Code for input of combined heat curve (Table 3, Vapour, Density)
#-------------------------------------------------------------------------------------------------------------------
for x in range (0, y):
    a = sheet.cell(row = x + 8, column = 17).value
    b = sheet.cell(row = x + 55, column = 17).value
    if x < (y-1)/2:
        sheet['AC' + str(x + 8)].value = a
    else:
        sheet['AC' + str(x + 8)].value = b
#-------------------------------------------------------------------------------------------------------------------



#Code for input of combined heat curve (Table 3, Vapour, Heat Capacity)
#-------------------------------------------------------------------------------------------------------------------
for x in range (0, y):
    a = sheet.cell(row = x + 8, column = 18).value
    b = sheet.cell(row = x + 55, column = 18).value
    if x < (y-1)/2:
        sheet['AD' + str(x + 8)].value = a
    else:
        sheet['AD' + str(x + 8)].value = b
#-------------------------------------------------------------------------------------------------------------------




#Code for input of combined heat curve (Table 3, Vapour, Viscocity)
#-------------------------------------------------------------------------------------------------------------------
for x in range (0, y):
    a = sheet.cell(row = x + 8, column = 19).value
    b = sheet.cell(row = x + 55, column = 19).value
    if x < (y-1)/2:
        sheet['AE' + str(x + 8)].value = a
    else:
        sheet['AE' + str(x + 8)].value = b
#-------------------------------------------------------------------------------------------------------------------



#Code for input of combined heat curve (Table 3, Vapour, Thermal Conductivity)
#-------------------------------------------------------------------------------------------------------------------
for x in range (0, y):
    a = sheet.cell(row = x + 8, column = 20).value
    b = sheet.cell(row = x + 55, column = 20).value
    if x < (y-1)/2:
        sheet['AF' + str(x + 8)].value = a
    else:
        sheet['AF' + str(x + 8)].value = b
#-------------------------------------------------------------------------------------------------------------------



#Code for input of combined heat curve (Table 3, Liquid, Density)
#-------------------------------------------------------------------------------------------------------------------
for x in range (0, y):
    a = sheet.cell(row = x + 8, column = 21).value
    b = sheet.cell(row = x + 55, column = 21).value
    if x < (y-1)/2:
        sheet['AG' + str(x + 8)].value = a
    else:
        sheet['AG' + str(x + 8)].value = b
#-------------------------------------------------------------------------------------------------------------------




#Code for input of combined heat curve (Table 3, Liquid, Heat Capacity)
#-------------------------------------------------------------------------------------------------------------------
for x in range (0, y):
    a = sheet.cell(row = x + 8, column = 22).value
    b = sheet.cell(row = x + 55, column = 22).value
    if x < (y-1)/2:
        sheet['AH' + str(x + 8)].value = a
    else:
        sheet['AH' + str(x + 8)].value = b
#-------------------------------------------------------------------------------------------------------------------




#Code for input of combined heat curve (Table 3, Liquid, Viscocity)
#-------------------------------------------------------------------------------------------------------------------
for x in range (0, y):
    a = sheet.cell(row = x + 8, column = 23).value
    b = sheet.cell(row = x + 55, column = 23).value
    if x < (y-1)/2:
        sheet['AI' + str(x + 8)].value = a
    else:
        sheet['AI' + str(x + 8)].value = b
#-------------------------------------------------------------------------------------------------------------------



#Code for input of combined heat curve (Table 3, Liquid, Thermal Conductivity)
#-------------------------------------------------------------------------------------------------------------------
for x in range (0, y):
    a = sheet.cell(row = x + 8, column = 24).value
    b = sheet.cell(row = x + 55, column = 24).value
    if x < (y-1)/2:
        sheet['AJ' + str(x + 8)].value = a
    else:
        sheet['AJ' + str(x + 8)].value = b
#-------------------------------------------------------------------------------------------------------------------




# << CHECK << flag for interpolating with values of 0 
#-------------------------------------------------------------------------------------------------------------------
#Table 2
#-------
for x in range(55, (54 + y)):
    for x1 in range (5, 12):
        a = sheet.cell(row = x, column = x1).value
        b = sheet.cell(row = x + 1, column = x1).value
        c = sheet.cell(row = x, column = 2).value
        d = sheet.cell(row = x + 1, column = 2).value
        if a == 0 and b != 0:
            for x2 in range (55, (55 + y)):
                e = sheet.cell(row = x2, column = 14).value
                if c > e > d:
                    sheet['Y' + str(x2)].value = '<< CHECK <<'
                else:
                    ''
        else:
            ''
for x in range(55, (54 + y)):
    for x1 in range (5, 12):
        a = sheet.cell(row = x, column = x1).value
        b = sheet.cell(row = x + 1, column = x1).value
        c = sheet.cell(row = x, column = 2).value
        d = sheet.cell(row = x + 1, column = 2).value
        if a != 0 and b == 0:
            for x2 in range (55, (55 + y)):
                e = sheet.cell(row = x2, column = 14).value
                if c > e > d:
                    sheet['Y' + str(x2)].value = '<< CHECK <<'
                else:
                    ''
        else:
            ''
#-------
#Table 1
#-------
for x in range(8, (7 + y)):
    for x1 in range (5, 12):
        a = sheet.cell(row = x, column = x1).value
        b = sheet.cell(row = x + 1, column = x1).value
        c = sheet.cell(row = x, column = 2).value
        d = sheet.cell(row = x + 1, column = 2).value
        if a == 0 and b != 0:
            for x2 in range (8, (8 + y)):
                e = sheet.cell(row = x2, column = 14).value
                if c > e > d:
                    sheet['Y' + str(x2)].value = '<< CHECK <<'
                else:
                    ''
        else:
            ''
for x in range(8, (7 + y)):
    for x1 in range (5, 12):
        a = sheet.cell(row = x, column = x1).value
        b = sheet.cell(row = x + 1, column = x1).value
        c = sheet.cell(row = x, column = 2).value
        d = sheet.cell(row = x + 1, column = 2).value
        if a != 0 and b == 0:
            for x2 in range (8, (8 + y)):
                e = sheet.cell(row = x2, column = 14).value
                if c > e > d:
                    sheet['Y' + str(x2)].value = '<< CHECK <<'
                else:
                    ''
        else:
            ''
#-------
#Table 3
#-------
for x in range (8, 8 + int(((y - 1)/2) + 1)):
    a = sheet.cell(row = x, column = 25).value
    if a == '<< CHECK <<':
        sheet['AK' + str(x)].value = '<< CHECK <<'
    else:
        ''

for x in range (55 + int(((y - 1)/2) + 1), 55 + y):
    a = sheet.cell(row = x, column = 25).value
    if a == '<< CHECK <<':
        print('')
        sheet['AK' + str(x - 47)].value = '<< CHECK <<'
    else:
        ''
#-------------------------------------------------------------------------------------------------------------------


#----------- SHEET 2 ---------------------------


#-------------------------------------------------------------------------------------------------------------------
# Code for 2nd Sheet, If ENTHALPY is inputted
#-------------------------------------------------------------------------------------------------------------------
sheet = wb['Sheet1']


#Code for 2nd Sheet, Table 3, Enthalpy
#-------------------------------------------------------------------------------------------------------------------
if sheet['B1'].value == 'Enthalpy':
    a = sheet['B4'].value
    b = sheet['C8'].value
    c = sheet['C9'].value
    sheet['P8'].value = a*b - a*c
    sheet['P9'].value = a*c - a*c
    for i in range(8, 8 + y):
        sheet = wb['Sheet']
        e = sheet.cell(row = i, column = 27).value
        sheet = wb['Sheet1']
        sheet['AD' + str(i)].value = e
        
# Code for 2nd Sheet Table 3, Temperature
#------------------------------------------------------------------------------------------

    x1 = sheet['B8'].value
    x2 = sheet['B9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AC' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, Vapour Fraction
#------------------------------------------------------------------------------------------
    x1 = sheet['D8'].value
    x2 = sheet['D9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AE' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, Vapour, Density
#------------------------------------------------------------------------------------------
    x1 = sheet['E8'].value
    x2 = sheet['E9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AF' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, Vapour, Heat Capacity
#------------------------------------------------------------------------------------------
    x1 = sheet['F8'].value
    x2 = sheet['F9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AG' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, Vapour, Viscocity
#------------------------------------------------------------------------------------------
    x1 = sheet['G8'].value
    x2 = sheet['G9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AH' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, Vapour, Viscocity
#------------------------------------------------------------------------------------------
    x1 = sheet['H8'].value
    x2 = sheet['H9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AI' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, liquid, Density
#------------------------------------------------------------------------------------------
    x1 = sheet['I8'].value
    x2 = sheet['I9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AJ' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, liquid, Heat Capacity
#------------------------------------------------------------------------------------------
    x1 = sheet['J8'].value
    x2 = sheet['J9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AK' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, liquid, Viscocity
#------------------------------------------------------------------------------------------
    x1 = sheet['K8'].value
    x2 = sheet['K9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AL' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, liquid, Thermal Conductivity
#------------------------------------------------------------------------------------------
    x1 = sheet['L8'].value
    x2 = sheet['L9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AM' + str(i)].value = gradient*a + intercept



#-------------------------------------------------------------------------------------------------------------------
# Code for 2nd Sheet, If DUTY is inputted
#-------------------------------------------------------------------------------------------------------------------

#Code for 2nd Sheet, Table 3, Enthalpy or Duty
#-------------------------------------------------------------------------------------------------------------------
elif sheet['B1'].value == 'Duty':
    for i in range(8, 8 + y):
        sheet = wb['Sheet']
        e = sheet.cell(row = i, column = 27).value
        sheet = wb['Sheet1']
        sheet['AD' + str(i)].value = e

# Code for 2nd Sheet Table 3, Temperature
#------------------------------------------------------------------------------------------

    x1 = sheet['B8'].value
    x2 = sheet['B9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AC' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, Vapour Fraction
#------------------------------------------------------------------------------------------
    x1 = sheet['D8'].value
    x2 = sheet['D9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AE' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, Vapour, Density
#------------------------------------------------------------------------------------------
    x1 = sheet['E8'].value
    x2 = sheet['E9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AF' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, Vapour, Heat Capacity
#------------------------------------------------------------------------------------------
    x1 = sheet['F8'].value
    x2 = sheet['F9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AG' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, Vapour, Viscocity
#------------------------------------------------------------------------------------------
    x1 = sheet['G8'].value
    x2 = sheet['G9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AH' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, Vapour, Viscocity
#------------------------------------------------------------------------------------------
    x1 = sheet['H8'].value
    x2 = sheet['H9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AI' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, liquid, Density
#------------------------------------------------------------------------------------------
    x1 = sheet['I8'].value
    x2 = sheet['I9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AJ' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, liquid, Heat Capacity
#------------------------------------------------------------------------------------------
    x1 = sheet['J8'].value
    x2 = sheet['J9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AK' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, liquid, Viscocity
#------------------------------------------------------------------------------------------
    x1 = sheet['K8'].value
    x2 = sheet['K9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AL' + str(i)].value = gradient*a + intercept

# Code for 2nd Sheet Table 3, liquid, Thermal Conductivity
#------------------------------------------------------------------------------------------
    x1 = sheet['L8'].value
    x2 = sheet['L9'].value
    y1 = sheet['AD8'].value
    y2 = sheet['AD' + str(7 + y)].value
    gradient = ((x2 - x1)/(y2 - y1))
    intercept = x1 - gradient*y1
    for i in range(8, 8 + y):
        a = sheet.cell(row = i, column = 30).value
        sheet['AM' + str(i)].value = gradient*a + intercept
        
else:
    ''

#SHEET 3
#------------------------------------------------------------------------------------------
sheet = wb['ZONAL INPUT']
for i in range(5, 50):
    sheet['B' + str(i)].value = ' '
    sheet['C' + str(i)].value = ' '
    sheet['D' + str(i)].value = ' '
    sheet['E' + str(i)].value = ' '
    sheet['F' + str(i)].value = ' '
    sheet['G' + str(i)].value = ' '
    sheet['H' + str(i)].value = ' '
    sheet['I' + str(i)].value = ' '
    sheet['J' + str(i)].value = ' '
    sheet['K' + str(i)].value = ' '
    sheet['L' + str(i)].value = ' '


for i in range(8, 8 + y):
    sheet = wb['Sheet']
    a = sheet.cell(row = i, column = 26).value
    sheet = wb['ZONAL INPUT']
    sheet['B' + str(i - 3)].value = a

for i in range(8, 8 + y):
    sheet = wb['Sheet']
    a = sheet.cell(row = i, column = 27).value
    sheet = wb['ZONAL INPUT']
    sheet['C' + str(i - 3)].value = a
    
for i in range(8, 8 + y):
    sheet = wb['Sheet']
    a = sheet.cell(row = i, column = 28).value
    sheet = wb['ZONAL INPUT']
    sheet['D' + str(i - 3)].value = a
    
for i in range(8, 8 + y):
    sheet = wb['Sheet']
    a = sheet.cell(row = i, column = 29).value
    sheet = wb['ZONAL INPUT']
    sheet['E' + str(i - 3)].value = a
    if a == 0:
        sheet['E' + str(i - 3)].value = '-'
    else:
        ''
        
for i in range(8, 8 + y):
    sheet = wb['Sheet']
    a = sheet.cell(row = i, column = 30).value
    sheet = wb['ZONAL INPUT']
    sheet['F' + str(i - 3)].value = a
    if a == 0:
        sheet['F' + str(i - 3)].value = '-'
    else:
        ''
    
for i in range(8, 8 + y):
    sheet = wb['Sheet']
    a = sheet.cell(row = i, column = 31).value
    sheet = wb['ZONAL INPUT']
    sheet['G' + str(i - 3)].value = a
    if a == 0:
        sheet['G' + str(i - 3)].value = '-'
    else:
        ''
    
for i in range(8, 8 + y):
    sheet = wb['Sheet']
    a = sheet.cell(row = i, column = 32).value
    sheet = wb['ZONAL INPUT']
    sheet['H' + str(i - 3)].value = a
    if a == 0:
        sheet['H' + str(i - 3)].value = '-'
    else:
        ''
    
for i in range(8, 8 + y):
    sheet = wb['Sheet']
    a = sheet.cell(row = i, column = 33).value
    sheet = wb['ZONAL INPUT']
    sheet['I' + str(i - 3)].value = a
    if a == 0:
        sheet['I' + str(i - 3)].value = '-'
    else:
        ''

for i in range(8, 8 + y):
    sheet = wb['Sheet']
    a = sheet.cell(row = i, column = 34).value
    sheet = wb['ZONAL INPUT']
    sheet['J' + str(i - 3)].value = a
    if a == 0:
        sheet['J' + str(i - 3)].value = '-'
    else:
        ''
    
for i in range(8, 8 + y):
    sheet = wb['Sheet']
    a = sheet.cell(row = i, column = 35).value
    sheet = wb['ZONAL INPUT']
    sheet['K' + str(i - 3)].value = a
    if a == 0:
        sheet['K' + str(i - 3)].value = '-'
    else:
        ''
    
for i in range(8, 8 + y):
    sheet = wb['Sheet']
    a = sheet.cell(row = i, column = 36).value
    sheet = wb['ZONAL INPUT']
    sheet['L' + str(i - 3)].value = a
    if a == 0:
        sheet['L' + str(i - 3)].value = '-'
    else:
        ''

for i in range(8, 8 + y):
    sheet = wb['Sheet']
    a = sheet.cell(row = i, column = 37).value
    sheet = wb['ZONAL INPUT']
    sheet['M' + str(i - 3)].value = a


#------------------------------------------------------------------------------------------
sheet = wb['ZONAL INPUT']
sheet['P1'].value = ' '


#calculation for non process mass flow
sheet = wb['Sheet1']
if sheet['B1'].value == 'Enthalpy' or 'Duty':
    E1 = sheet.cell(row = 8, column = 3).value
    E2 = sheet.cell(row = 9, column = 3).value
    D1 = sheet.cell(row = 8, column = 30).value
    if isinstance(E1, (float, int)) and isinstance(E2, (float, int)) and isinstance(D1, (float, int)):
        sheet = wb['ZONAL INPUT']
        sheet['P1'].value = ((D1)/(E1-E2))
    else:
        ''
else:
    ''


    


for i in range(8, 8 + y):
    sheet = wb['Sheet1']
    a = sheet.cell(row = i, column = 29).value
    sheet = wb['ZONAL INPUT']
    sheet['O' + str(i - 3)].value = a

for i in range(8, 8 + y):
    sheet = wb['Sheet1']
    a = sheet.cell(row = i, column = 30).value
    sheet = wb['ZONAL INPUT']
    sheet['P' + str(i - 3)].value = a
    
for i in range(8, 8 + y):
    sheet = wb['Sheet1']
    a = sheet.cell(row = i, column = 31).value
    sheet = wb['ZONAL INPUT']
    sheet['Q' + str(i - 3)].value = a
    
for i in range(8, 8 + y):
    sheet = wb['Sheet1']
    a = sheet.cell(row = i, column = 32).value
    sheet = wb['ZONAL INPUT']
    sheet['R' + str(i - 3)].value = a
    if a == 0:
        sheet['R' + str(i - 3)].value = '-'
    else:
        ''
        
for i in range(8, 8 + y):
    sheet = wb['Sheet1']
    a = sheet.cell(row = i, column = 33).value
    sheet = wb['ZONAL INPUT']
    sheet['S' + str(i - 3)].value = a
    if a == 0:
        sheet['S' + str(i - 3)].value = '-'
    else:
        ''
    
for i in range(8, 8 + y):
    sheet = wb['Sheet1']
    a = sheet.cell(row = i, column = 34).value
    sheet = wb['ZONAL INPUT']
    sheet['T' + str(i - 3)].value = a
    if a == 0:
        sheet['T' + str(i - 3)].value = '-'
    else:
        ''
    
for i in range(8, 8 + y):
    sheet = wb['Sheet1']
    a = sheet.cell(row = i, column = 35).value
    sheet = wb['ZONAL INPUT']
    sheet['U' + str(i - 3)].value = a
    if a == 0:
        sheet['U' + str(i - 3)].value = '-'
    else:
        ''
    
for i in range(8, 8 + y):
    sheet = wb['Sheet1']
    a = sheet.cell(row = i, column = 36).value
    sheet = wb['ZONAL INPUT']
    sheet['V' + str(i - 3)].value = a
    if a == 0:
        sheet['V' + str(i - 3)].value = '-'
    else:
        ''

for i in range(8, 8 + y):
    sheet = wb['Sheet1']
    a = sheet.cell(row = i, column = 37).value
    sheet = wb['ZONAL INPUT']
    sheet['W' + str(i - 3)].value = a
    if a == 0:
        sheet['W' + str(i - 3)].value = '-'
    else:
        ''
    
for i in range(8, 8 + y):
    sheet = wb['Sheet1']
    a = sheet.cell(row = i, column = 38).value
    sheet = wb['ZONAL INPUT']
    sheet['X' + str(i - 3)].value = a
    if a == 0:
        sheet['X' + str(i - 3)].value = '-'
    else:
        ''
    
for i in range(8, 8 + y):
    sheet = wb['Sheet1']
    a = sheet.cell(row = i, column = 39).value
    sheet = wb['ZONAL INPUT']
    sheet['Y' + str(i - 3)].value = a
    if a == 0:
        sheet['Y' + str(i - 3)].value = '-'
    else:
        ''

for i in range(8, 8 + y):
    sheet = wb['Sheet1']
    a = sheet.cell(row = i, column = 40).value
    sheet = wb['ZONAL INPUT']
    sheet['Z' + str(i - 3)].value = a









print('Calculations completed successfully')


wb.save('ZIP.xlsx')
