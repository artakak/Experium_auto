# -*- coding: utf-8 -*-
import win32con
from pywinauto.application import Application
from faker import Factory
import time


"""pywinauto.findwindows.find_elements(class_name=None, class_name_re=None, parent=None, process=None, title=None,
title_re=None, top_level_only=True, visible_only=True, enabled_only=False, best_match=None, handle=None,
ctrl_index=None, found_index=None, predicate_func=None, active_only=False, control_id=None, control_type=None,
auto_id=None, framework_id=None, backend=None)"""

fake = Factory.create('ru_RU')
app = Application(backend="win32").connect(title_re="Experium")
# app2 = Application(backend="uia").connect(title_re="Experium")
dlg = app.top_window()
dlg.set_focus()

# New_people_card
# dlg.child_window(control_id=30210).send_message_timeout(0x0400+0x4024, 200, 0)
dlg.child_window(control_id=30210).click()
popup = app.top_window()
# print(popup.menu().items())
popup.menu().Item(4).click_input()
time.sleep(3)
Chel_card = app.top_window()
# Chel_card.print_control_identifiers()

# People_edit_mainInfo
main_info = Chel_card['PeopleEditMainInfo']
#main_info.print_control_identifiers()
print(main_info.child_window(control_id=2422).SetText(fake.last_name_male()).window_text())
main_info.child_window(control_id=2423).SetText(fake.first_name_male())
main_info.child_window(control_id=2424).SetText(fake.first_name_male())
main_info.child_window(control_id=23712).SetText(fake.last_name_male())
main_info.child_window(control_id=2425).send_chars('12121988')
main_info.child_window(control_id=2426).set_focus().type_keys('{DOWN 2}{ENTER}')
# print(Chel_card.child_window(control_id=2426).window_text())
print(main_info.child_window(control_id=4509).set_focus().type_keys('{DOWN 10}{ENTER}').window_text())
# print(Chel_card['Edit6'].window_text())

# People_edit_personalInfo
maininfo_2 = Chel_card['PeopleEditPersonalInfo']
maininfo_2.child_window(control_id=100).send_message_timeout(win32con.WM_SETTEXT, 0, '4')
maininfo_2.child_window(control_id=101).send_message_timeout(win32con.WM_SETTEXT, 0, '9')
maininfo_2.child_window(control_id=102).send_message_timeout(win32con.WM_SETTEXT, 0, 'B')
maininfo_2.child_window(control_id=103).set_focus().type_keys('{DOWN 2}{ENTER}')
maininfo_2.child_window(control_id=104).set_focus().type_keys('{DOWN 2}{ENTER}')
print(Chel_card.child_window(control_id=104).window_text())

# SalaryExpectations
for n in range(1):
    Chel_card.child_window(control_id=23849).SetText('10000' + str(n))
    Chel_card['Edit8'].SetText('EUR')
    Chel_card.child_window(control_id=23848).set_focus()
    Chel_card.child_window(control_id=23848).send_chars('1212201' + str(n))
    Chel_card.child_window(control_id=2422).click_input()
    Chel_card.child_window(control_id=23847).SetText('Text5_' + str(n))
    Chel_card['+'].click_input()

# Position
for n in range(1):
    Chel_card.child_window(control_id=24003).SetText('Place_' + str(n))
    Chel_card.child_window(control_id=24005).set_focus().send_chars('12201' + str(n))
    Chel_card.child_window(control_id=2422).click_input()
    Chel_card.child_window(control_id=24006).set_focus().send_chars('12202' + str(n))
    Chel_card.child_window(control_id=2422).click_input()
    Chel_card.child_window(control_id=24004).SetText('Dol_' + str(n))
    Chel_card.child_window(control_id=24007).SetText('Otdel_' + str(n))
    Chel_card['+2'].click_input()

# Position_fromDB
Chel_card.child_window(control_id=32019).click_input()
Chel_card['Edit11'].set_focus().type_keys('{DOWN 2}{ENTER}')
Chel_card.child_window(control_id=24005).set_focus().send_chars('122010')
Chel_card.child_window(control_id=2422).click_input()
Chel_card.child_window(control_id=24006).set_focus().send_chars('122020')
Chel_card.child_window(control_id=2422).click_input()
Chel_card.child_window(control_id=24004).SetText('Dol_DB')
Chel_card['+2'].click_input()

Chel_card.child_window(control_id=803).click_input()


# time.sleep(3)
# confirm = app.top_window()
# confirm.child_window(control_id=100).click_input()
# listbox = Chel_card.child_window(control_id=1007)
# print(listbox.wrapper_object().menu_items())
