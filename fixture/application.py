import clr
import os.path

project_dir = os.path.dirname(os.path.abspath(__file__))
import sys

sys.path.append(os.path.join(project_dir, "TestStack.White.0.13.3\\lib\\net40\\"))
sys.path.append(os.path.join(project_dir, "Castle.Core.3.3.0\\lib\\net40-client\\"))
clr.AddReferenceByName('TestStack.White')

from TestStack.White import Application
from TestStack.White.UIItems.Finders import *

clr.AddReferenceByName('UIAutomationTypes, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
from System.Windows.Automation import *

from fixture.group import GroupHelper



class Application:
    def __init__(self, path, name_window):
        self.app = Application.Launch(path)
        self.main_window = self.app.GetWindow(name_window)
        self.group = GroupHelper(self)

    def destroy(self):
        self.app.Get(SearchCriteria.ByAutomationId("uxExitAddressButton")).Click()



class Application

    def __init__(self):
        self.appl =

    def open_application(self):
        application = Application.Launch("c:\\Downloads\\Addressbook\\AddressBook.exe")