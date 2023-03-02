import xlwings as xw
import os

def test_isDeveloper():
    os.environ["TEST"] = "dev"
    wb = xw.Book('VBAWorkbook.xlsm')

    getEnvironment = wb.macro('GetEnvironment')
    
    assert os.environ["TEST"] == "dev"
    assert getEnvironment() == "dev"