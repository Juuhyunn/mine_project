import xlwings as xw


def main():
    # add()
    vlookup()


def add():
    a = xw.Book.caller()
    sheet = a.sheets[0]
    i = 1
    sheet[f'A{i}'].value = i
    for i in range(4):
        sheet[f'C{i+1}'].value = sheet[f'A{i+1}'].value + sheet[f'B{i+1}'].value


def vlookup():
    a = xw.Book.caller()
    sheet = a.sheets[0]
    for i in range(10):
        if sheet[f'A{i+1}'].value == sheet[f'B{i+1}'].value:
            sheet[f'C{i+1}'].value = True
        else:
            sheet[f'C{i+1}'].value = False






if __name__ == '__main__':
    xw.Book("hello.xlsm").set_mock_caller()
    main()