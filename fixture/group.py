import clr
import os.path
project_dir = os.path.dirname(os.path.abspath(__file__))
import sys
sys.path.append(os.path.join(project_dir, "TestStack.White.0.13.3\\lib\\net40\\"))
sys.path.append(os.path.join(project_dir, "Castle.Core.3.3.0\\lib\\net40-client\\"))
clr.AddReferenceByName('TestStack.White')

from TestStack.White.InputDevices import Keyboard
from TestStack.White.WindowsAPI import KeyboardInput
from TestStack.White.UIItems.Finders import *

clr.AddReferenceByName('UIAutomationTypes, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
from System.Windows.Automation import *

class GroupHelper:

    def __init__(self, app):
        self.app = app

    def open_group_editor(self):
        self.app.main_window.Get(SearchCriteria.ByAutomationId("groupButton")).Click()
        modal = self.app.main_window.ModalWindow("Group editor")
        return modal

    def close_group_editor(self, modal):
        modal.Get(SearchCriteria.ByAutomationId("uxCloseAddressButton")).Click()

    def add_new_group(self, name):
        modal = self.open_group_editor()
        modal.Get(SearchCriteria.ByAutomationId("uxNewAddressButton")).Click()
        modal.Get(SearchCriteria.ByControlType(ControlType.Edit)).Enter(name)
        Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RETURN)
        self.close_group_editor(modal)

    def delete_group(self, group):
        modal = self.open_group_editor()
        modal.Get(SearchCriteria.ByText("%s" % group)).Click()
        modal.Get(SearchCriteria.ByAutomationId("uxDeleteAddressButton")).Click()
        modal.Get(SearchCriteria.ByAutomationId("uxDeleteAllRadioButton")).Click()
        modal.Get(SearchCriteria.ByAutomationId("uxOKAddressButton")).Click()

        self.close_group_editor(modal)

    def get_group_list(self):
        modal = self.open_group_editor()
        tree = modal.Get(SearchCriteria.ByAutomationId("uxAddressTreeView"))
        root = tree.Nodes[0]
        l = [node.Text for node in root.Nodes]
        self.close_group_editor(modal)
        return l