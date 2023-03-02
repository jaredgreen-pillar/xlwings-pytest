import xlwings as xw
import os

def test_GetEnvironment():
    os.environ["TEST"] = "dev"
    wb = xw.Book('VBAWorkbook.xlsm')

    getEnvironment = wb.macro('GetEnvironment')
    
    assert os.environ["TEST"] == "dev"    # Passes - the env variable is set
    assert getEnvironment() == "dev"      # Fails - VBA code cannot access the "TEST" env variable