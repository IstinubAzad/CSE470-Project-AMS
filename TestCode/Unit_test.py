import unittest
import sys

sys.path.append("..")
import TestCode.Test_app as tApp


class TestMethods(unittest.TestCase):

    def test_eBill_total(self):
        sum_from_sheet = 100 + 200 + 300
        TestMethods().assertEqual(sum_from_sheet, tApp.Bill_Calculation().electric_Bill())
        print('DONE --- test for total sum of electric bill matching with excel sheet values')

    def test_wBill_total(self):
        sum_from_sheet = 100 + 200 + 300
        TestMethods().assertEqual(sum_from_sheet, tApp.Bill_Calculation().water_bill())
        print('DONE --- test for total sum of water bill matching with excel sheet values')

    def test_gBill_total(self):
        sum_from_sheet = 100 + 200 + 300
        TestMethods().assertEqual(sum_from_sheet, tApp.Bill_Calculation().gas_Bill())
        print('DONE --- test for total sum of gas bill matching with excel sheet values')

    def test_MonthlyBill_Flat01(self):
        listValues = tApp.Bill_database_read().load_workbook_values()
        row = 1
        sumEachRow = 0
        for i in range(3,6):
            sumEachRow += int(listValues[row][i])
        TestMethods().assertEqual(sumEachRow, int(listValues[row][6]))
        print('DONE --- Total monthly bill of each flat IN 1ST row')

    def test_MonthlyBill_Flat02(self):
        listValues = tApp.Bill_database_read().load_workbook_values()
        row = 2
        sumEachRow = 0
        for i in range(3,6):
            sumEachRow += int(listValues[row][i])
        TestMethods().assertEqual(sumEachRow, int(listValues[row][6]))
        print('DONE --- Total monthly bill of each flat 2ND row')

    def test_MonthlyBill_Flat03(self):
        listValues = tApp.Bill_database_read().load_workbook_values()
        row = 3
        sumEachRow = 0
        for i in range(3,6):
            sumEachRow += int(listValues[row][i])
        TestMethods().assertEqual(sumEachRow, int(listValues[row][6]))
        print('DONE --- test for total monthly bill of each flat 3RD row')

    def test_total_monthly_bill(self):
        value_from_Sheet = tApp.Bill_Calculation().water_bill() + tApp.Bill_Calculation().electric_Bill() + tApp.Bill_Calculation().gas_Bill()
        total_value_from_code = tApp.Bill_Calculation().total_monthly_bill()
        TestMethods().assertEqual(value_from_Sheet, total_value_from_code)
        print('DONE --- test for total monthly bill of all flats')

    def test_flatList_BillList(self):
        Bill_List = tApp.Bill_database_read().load_workbook_values()
        Flat_List = tApp.Flat_database_read().load_workbook_values()
        TestMethods().assertEqual(len(Flat_List), len(Bill_List))
        print('DONE --- test for matching the total number of rows in bill database and row database')
