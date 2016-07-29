

class GroupHelper:

    def __init__(self, app):
        self.app = app

    def open_group_editor(main_window):
        main_window.Get(SearchCriteria.ByAutomationId("groupButton")).Click()
        modal = main_window.ModalWindow("Group editor")
        return modal


    def close_group_editor(modal):
        modal.Get(SearchCriteria.ByAutomationId("uxCloseAddressButton")).Click()


    def add_new_group(main_window, name):
        modal = open_group_editor(main_window)
        modal.Get(SearchCriteria.ByAutomationId("uxNewAddressButton")).Click()
        modal.Get(SearchCriteria.ByControlType(ControlType.Edit)).Enter(name)
        Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RETURN)
        close_group_editor(modal)

    def get_group_list(main_window):
        modal = open_group_editor(main_window)
        tree = modal.Get(SearchCriteria.ByAutomationId("uxAddressTreeView"))
        root = tree.Nodes[0]
        l = [node.Text for node in root.Nodes]
        close_group_editor(modal)
        return l

    def delete_group(main_window, name):
        modal = open_group_editor(main_window)
