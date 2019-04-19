import os
import xlsxwriter

from bs4 import BeautifulSoup


def get_data():

    a = []
    file = "C:\\Users\\zzheng\\Documents\\Logs\\IntekCTS\\Result\\test_result_failures.html"
    soup = BeautifulSoup(open(file), "html.parser")
    summary = soup.find('table', attrs={'class': 'testsummary'})
    trs = summary.find_all('tr')

    for tr in trs:
        tds = tr.find_all('td')

        b = []
        for td in tds:
            content = td.text.strip()
            b.append(content)
        a.append(b)
    return a


def generate_excel(data):
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
    bold = workbook.add_format({'bold': True})
    row = 1
    sheet.write('A1', "Module Case", bold)
    sheet.write('B1', "Passed Number", bold)
    sheet.write('C1', "Failed Number", bold)
    sheet.write('D1', "Module Completed", bold)

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
    data = get_data()
    generate_excel(data)
