import os.path
import clr
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

project_dir = os.path.dirname(os.path.abspath(__file__))
import sys
sys.path.append(os.path.join(project_dir, "TestStack.White.0.13.3\\lib\\net40\\"))
sys.path.append(os.path.join(project_dir, "Castle.Core.3.3.0\\lib\\net40-client\\"))
clr.AddReferenceByName('TestStack.White')

from TestStack.White.UIItems.Finders import *

import json
import pytest
from fixture.application import Application_Window

fixture = None
conf = None

def load_config(file):
    global conf
    if conf is None:
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), file)) as json_file:
            conf = json.load(json_file)
    return conf

@pytest.fixture
def app(request):
    global fixture
    app_conf = load_config("config.json")
    if fixture is None:
        fixture = Application_Window(path=app_conf["application"]["path"], window=app_conf["application"]["main_window"])
    return fixture

@pytest.fixture(scope="session", autouse=True)
def stop(request):
    def exit():
        fixture.destroy()
    request.addfinalizer(exit)
    return fixture

def pytest_generate_tests(metafunc):
    for fixture in metafunc.fixturenames:
        if fixture.startswith("data_"):
            testdata = load_from_excel(fixture[5:])
            metafunc.parametrize(fixture, testdata, ids=[str(x) for x in testdata])

def load_from_excel(file):
    excel = Excel.ApplicationClass()
    workbook = excel.Workbooks.Open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data\%s.xlsx" % file))
    worksheet = workbook.ActiveSheet
    data = []
    i = 1
    while worksheet.Range["A%s" % str(i)].Value2 is not None:
        data.append(worksheet.Range["A%s" % str(i)].Value2)
        i = i + 1
    excel.Quit()
    return data