import openpyxl
import queue
import pyautogui, time

filename = '/Users/du/Desktop/dd/2022年1月份数据/2022年1月10日.xlsx'
file = openpyxl.load_workbook(filename)
sheets = file.get_sheet_names()
sheet = file.get_sheet_by_name('Sheet1')
name = ''
phone_number = ''



q = queue.Queue(1000)
for i in range(58):
    a1 = 'F' + str(3 * (i + 1))
    a2 = 'F' + str(3 * (i + 1) + 1)
    a3 = 'F' + str(3 * (i + 1) + 2)
    a4 = 'G' + str(3 * (i + 1))
    a5 = 'I' + str(3 * (i + 1))
    a6 = 'I' + str(3 * (i + 1) + 1)
    a7 = 'I' + str(3 * (i + 1) + 2)
    a8 = 'J' + str(3 * (i + 1))
    q.put(sheet[a1].value)
    q.put(sheet[a2].value)
    q.put(sheet[a3].value)
    q.put(sheet[a4].value)
    q.put(sheet[a5].value)
    q.put(sheet[a6].value)
    q.put(sheet[a7].value)
    q.put(sheet[a8].value)
q.put(name)
q.put(phone_number)
time.sleep(5)
while(not q.empty()):
    pyautogui.typewrite(str(q.get()))
    pyautogui.press('tab')
    # print(str(q.get()) + '\n')




