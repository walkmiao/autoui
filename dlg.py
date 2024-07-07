import time
from typing import List
from pywinauto.keyboard import send_keys

def combox_select(sel, combo):
    """

    @param sel: 要选择的文本
    @param combo: combox控件
    @return:
    """
    combo.type_keys("{ENTER}")  # Selects the combo box
    texts = combo.texts()  # gets all texts available in combo box
    try:
        index = texts.index(str(sel))  # find index of required selection
    except ValueError:
        return False
    sel_index = combo.selected_index()  # find current index of combo
    if index > sel_index:
        combo.type_keys("{DOWN}" * abs(index - sel_index))
    else:
        combo.type_keys("{UP}" * abs(index - sel_index))


def dgv_data(dlg):
    """
    获取DataGridView控件的数据
    @param dlg:
    @return: 数据
    """
    result = []
    for row in dlg.children(control_type='Custom'):
        row_data = []
        result.append(row_data)
        for item in row.children(control_type="DataItem"):
            row_data.append(item.iface_value.CurrentValue)
    return result


def save_and_exit():
    send_keys('{F6}')
    time.sleep(1)
    send_keys('{F12}')