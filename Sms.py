from appJar import gui
import re
import xlrd
import clipboard
from datetime import timedelta
from datetime import datetime
from dateutil import  parser



#############################################################################

workbook = xlrd.open_workbook("warid.xlsx")
worksheet = workbook.sheet_by_index(0)

############################################################################

city = []
rbu = []
sites = []


############################################################################

for x in range(1, 5463):
    rbu.append(worksheet.cell(x, 17).value)
    city.append(worksheet.cell(x, 5).value)
    sites.append(worksheet.cell(x, 0).value)

for y in range(0, 5462):
    temp = sites[y]
    sites[y] = temp[-4:]

############################################################################



def press(btn):
    x = clipboard.paste()
    x = re.split(r'\t+', x)
    temp = x[4]
    temp = temp[:-2]
    x[4] = temp
    temp = temp[-4:]
    start_time = parser.parse(x[2])
    esc_time = start_time + timedelta(minutes=5)
    start_time = datetime.strftime(start_time, '%d/%m/%Y  %H:%M')
    esc_time = datetime.strftime(esc_time, '%d/%m/%Y  %H:%M')
    x[2] = str(start_time)
    esc_time = str(esc_time)


    for z in range(0, 5461):
        if temp == sites[z]:
            x.append(city[z])
            x.append(rbu[z])
        else:
            temp1 = 0

    y = x[4] + " is down " + "\r\n"  \
    + "City: " + x[5] + "-" + x[6] + "\r\n" \
    + "Reason: " + x[0] + "\r\n" \
    + "Start Time: " + x[2] + "\r\n" \
    + "Escalation Time: " + esc_time + "\r\n"  \
    + "Action Taken: " + "Escalated to NOSS and FM " + "\r\n" \
    + "TT no. : " + x[1]

    clipboard.copy(y)
#######################################################

def pressup(btn):
    x = clipboard.paste()
    x = re.split(r'\t+', x)
    temp = x[4]
    temp = temp[:-2]
    x[4] = temp
    temp = temp[-4:]
    start_time = parser.parse(x[2])
    esc_time = start_time + timedelta(minutes=5)
    start_time = datetime.strftime(start_time, '%d/%m/%Y  %H:%M')
    esc_time = datetime.strftime(esc_time, '%d/%m/%Y  %H:%M')
    x[2] = str(start_time)
    esc_time = str(esc_time)


    for z in range(0, 5461):
        if temp == sites[z]:
            x.append(city[z])
            x.append(rbu[z])
        else:
            temp1 = 0

    y = x[4] + " is down " + "\r\n"  \
    + "City: " + x[5] + "-" + x[6] + "\r\n" \
    + "Reason: " + x[0] + "\r\n" \
    + "Start Time: " + x[2] + "\r\n" \
    + "Escalation Time: " + esc_time + "\r\n"  \
    + "Action Taken: " + "Escalated to NOSS and FM " + "\r\n" \
    + "TT no. : " + x[1]

    clipboard.copy(y)

#######################################################

def press4Gup(btn):
    x = clipboard.paste()
    x = re.split(r'\t+', x)
    temp = x[4]
    temp = temp[:-2]
    x[4] = temp
    temp = temp[-4:]
    start_time = parser.parse(x[3])
    esc_time = start_time + timedelta(minutes=5)
    start_time = datetime.strftime(start_time, '%d/%m/%Y  %H:%M')
    esc_time = datetime.strftime(esc_time, '%d/%m/%Y  %H:%M')
    x[3] = str(start_time)
    esc_time = str(esc_time)


    for z in range(0, 5461):
        if temp == sites[z]:
            x.append(city[z])
            x.append(rbu[z])
        else:
            temp1 = 0

    y = "Severity = Major (Start)" + "\r\n" \
    +x[4] + " (4G/LTE) is down " + "\r\n"  \
    +"Impact: 4G/LTE services are down" + "\r\n" \
    + "City: " + x[5] + "-" + x[6] + "\r\n" \
    + "Reason: " + x[0] + "\r\n" \
    + "Start Time: " + x[3] + "\r\n" \
    + "Escalation Time: " + esc_time + "\r\n"  \
    + "Action Taken: " + "Escalated to NOSS and FM " + "\r\n" \
    + "TT no. : " + x[1]

    clipboard.copy(y)




#######################################################

def press4G(btn):
    x = clipboard.paste()
    x = re.split(r'\t+', x)
    temp = x[4]
    temp = temp[:-2]
    x[4] = temp
    temp = temp[-4:]
    start_time = parser.parse(x[3])
    esc_time = start_time + timedelta(minutes=5)
    start_time = datetime.strftime(start_time, '%d/%m/%Y  %H:%M')
    esc_time = datetime.strftime(esc_time, '%d/%m/%Y  %H:%M')
    x[3] = str(start_time)
    esc_time = str(esc_time)


    for z in range(0, 5461):
        if temp == sites[z]:
            x.append(city[z])
            x.append(rbu[z])
        else:
            temp1 = 0

    y = "Severity = Major (Start)" + "\r\n" \
    +x[4] + " (4G/LTE) is down " + "\r\n"  \
    +"Impact: 4G/LTE services are down" + "\r\n" \
    + "City: " + x[5] + "-" + x[6] + "\r\n" \
    + "Reason: " + x[0] + "\r\n" \
    + "Start Time: " + x[3] + "\r\n" \
    + "Escalation Time: " + esc_time + "\r\n"  \
    + "Action Taken: " + "Escalated to NOSS and FM " + "\r\n" \
    + "TT no. : " + x[1]

    clipboard.copy(y)



#######################################################

app = gui("SMS","200x300")

app.setSticky("news")

app.setExpand("both")

app.addButton("2G down SMS", press,0,0)

app.addButton("4G down SMS", press4G,1,0)

app.addButton("2G up SMS", pressup,0,2)

app.addButton("4G up SMS", press4Gup,1,2)

app.setButtonBg("2G down SMS", "red")

app.setButtonBg("2G up SMS", "green")

app.setButtonBg("4G down SMS", "red")

app.setButtonBg("4G up SMS", "green")

#######################################################
app.go()
