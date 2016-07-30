import random

def test_add_group(app, data_groups):
    old_list = app.group.get_group_list()
    app.group.add_new_group(data_groups)
    new_list = app.group.get_group_list()
    old_list.append(data_groups)
    assert sorted(old_list) == sorted(new_list)

def test_delete_group(app, data_groups):
    old_list = app.group.get_group_list()
    if len(old_list) == 2:
        app.group.add_new_group(data_groups)
    random_group = random.choice(old_list)
    app.group.delete_group(random_group)
    new_list = app.group.get_group_list()
    old_list.remove(random_group)
    assert sorted(old_list) == sorted(new_list)