from model.group import Group
import random
import string
import os.path
import getopt
import sys
import time

import clr
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

try:
    opts, args = getopt.getopt(sys.argv[1:], "n:f:", ["number of groups", "file"])
except getopt.GetoptError as err:
    getopt.usage()
    sys.exit(2)

number = 2
file = "data/groups.xlsx"

for o, a in opts:
    if o == "-n":
        number = int(a)
    elif o == "-f":
        file = a

def random_string(prefix, maxlen):
    char = string.ascii_letters + string.hexdigits + ' '*5 #+ string.punctuation
    return prefix + "".join([random.choice(char) for i in range(random.randrange(maxlen))])

random_data = [
    Group(name=random_string("name", 20))
    for i in range(number)
]

file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", file)

excel = Excel.ApplicationClass()
excel.Visible = True

workbook = excel.Workbooks.Add()
sheet = workbook.ActiveSheet

for i in range(len(random_data)):
    sheet.Range["A%s" % (i+1)].Value2 = random_data[i].name

workbook.SaveAs(file)

time.sleep(3)

excel.Quit()


