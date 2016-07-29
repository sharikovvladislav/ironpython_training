import clr
import os.path
import clr
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

import pytest
import json
import jsonpickle
from fixture.application import Application

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
    app_conf = load_config(request.config.getoption("--config"))['application']
    fixture = Application()
    fixture.session.start_application(app_conf["path"])
    return fixture

@pytest.fixture(scope="session", autouse=True)
def stop(request):
    def exit():
        fixture.destroy()
    request.addfinalizer(exit)
    return fixture

def pytest_addoption(parser):
    parser.addoption("--config", action="store", default="config.json")

def pytest_generate_tests(metafunc):
    for fixture in metafunc.fixturenames:
        if fixture.startswith("data_"):
            testdata = load_from_excel(fixture[5:])
            metafunc.parametrize(fixture, testdata, ids=[str(x) for x in testdata])

def load_from_excel(table):
    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data/%s.xlsx" % table)) as file:
        return jsonpickle.decode(file.read())
