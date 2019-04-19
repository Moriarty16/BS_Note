import os
import xlsxwriter

from bs4 import BeautifulSoup


def get_data():

    a = []
    width = [0, 0, 0, 0, 0]

    file = "C:\\Users\\zzheng\\Documents\\Logs\\IntekCTS\\Result\\test_result_failures.html"
    soup = BeautifulSoup(open(file), "html.parser")
    summary = soup.find('table', attrs={'class': 'testsummary'})
    trs = summary.find_all('tr')
    for tr in trs:
        tds = tr.find_all('td')
        if not tds:
            continue
        if len(tds) != 5:
            print("Length is not 5, is ", len(tds))
        b = []
        # print(tds)
        # for td in tds:
        #     content = td.text.strip()
        #     if len(content) > column01_longest_width:
        #         column01_longest_width = len(content)
        #     b.append(content)
        # a.append(b)
        for i in range(len(tds)):
            content = tds[i].text.strip()
            # print(content)
            # print("Element %s, length is: " % content, len(content))
            if len(content) > width[i]:
                width[i] = len(content)
                # print("The %d element is: " % i, width[i])
            # if len(tds[0]) > column00_longest_width:
            #     column00_longest_width = len(tds[0])
            #     print("column00_longest_width", column00_longest_width)
            # if len(tds[1]) > column01_longest_width:
            #     column01_longest_width = len(tds[1])
            #     print("column01_longest_width", column01_longest_width)
            # if len(tds[2]) > column02_longest_width:
            #     column02_longest_width = len(tds[2])
            #     print("column02_longest_width", column02_longest_width)
            # if len(tds[3]) > column03_longest_width:
            #     column03_longest_width = len(tds[3])
            #     print("column03_longest_width", column03_longest_width)
            # if len(tds[4]) > column04_longest_width:
            #     column04_longest_width = len(tds[4])
            #     print("column04_longest_width", column04_longest_width)

            b.append(content)
        a.append(b)
    return a, width


def generate_excel(data, width):
    print("Original width is: ", width)
    f = "Result.xlsx"

    try:
        print("File exists, remove it now.")
        os.remove(f)
    except FileNotFoundError:
        pass
    except PermissionError:
        raise RuntimeError("File exists, but no enough permission to delete it now.")

    workbook = xlsxwriter.Workbook("Result.xlsx")
    sheet = workbook.add_worksheet("Show")
    head00 = "Module Case"
    head01 = "Passed Number"
    head02 = "Failed Number"
    head03 = "Total Tests"
    head04 = "Module Completed"
    # new_width = [0, 0, 0, 0, 0]
    if len(head00) > width[0]:
        width[0] = len(head00)
        print("Re-assign 00")
    if len(head01) > width[1]:
        width[1] = len(head01)
        print("Re-assign 01")
    if len(head02) > width[2]:
        width[2] = len(head02)
        print("Re-assign 02")
    if len(head03) > width[3]:
        width[3] = len(head03)
        print("Re-assign 03")
    if len(head04) > width[4]:
        width[4] = len(head04)
        print("Re-assign 04")
    sheet.set_column(0, 0, width[0])
    sheet.set_column(1, 1, width[1])
    sheet.set_column(2, 2, width[2])
    sheet.set_column(3, 3, width[3])
    sheet.set_column(4, 4, width[4])
    print("New width is: ", width)
    bold = workbook.add_format({'bold': True})
    row = 0
    sheet.write('A1', head01, bold)
    sheet.write('B1', head02, bold)
    sheet.write('C1', head03, bold)
    sheet.write('D1', head04, bold)

    for i in data:
        col = 0
        for j in i:
            sheet.write(row, col, j)
            col = col + 1
        row = row + 1

    workbook.close()

    if os.path.exists(f):
        print(os.path.abspath(os.path.abspath(f)))
    else:
        print("Fail to create xlsx file.")


class Analyzer:
    data, width = get_data()
    generate_excel(data, width)
