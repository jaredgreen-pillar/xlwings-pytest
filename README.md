# xlwings-pytest
Sample test of VBA code with Python (via pytest and xlwings). The desire is to set environment variables as part of a test setup that are then utilized by the VBA code.

----

### Issue

The xlwings `wb.macro('...')()` function does not respect environment variables set via Python.

### Requirements
- pytest installed: `pip install -U pytest`

### Running tests
- `pytest`


