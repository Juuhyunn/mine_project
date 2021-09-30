import xlwings as xw


def main():
    a = xw.Book.caller()
    sheet = a.sheets[0]
    i = 1
    sheet[f'A{i}'].value = i

    # print(sheet[f'A{i}'].value + sheet['A2'].value)
    # if target.value == "Hello xlwings!" :
    #     target.value = "Bye xlwings!"
    # else :
    #     target.value = "Hello xlwings!"
    #
    for i in range(4):
        sheet[f'C{i+1}'].value = sheet[f'A{i+1}'].value + sheet[f'B{i+1}'].value




if __name__ == '__main__':
    xw.Book("hello.xlsm").set_mock_caller()
    main()