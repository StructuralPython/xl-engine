import xl_engine as xl
import math
import pathlib

TEST_DATA_DIR = pathlib.Path(__file__).parent / "test_data"

def test_create_condition_check():
    func = xl.create_condition_check(1.0, "le")
    assert func(2) == False
    assert func(1) == True
    assert func(0.9) == True

    func2 = xl.create_condition_check(2393.90, "gt")
    assert func2(3000) == True
    assert func2(1) == False
    assert func2(2393.90) == False

def test_execute_workbook():
    results = xl.execute_workbook(
        TEST_DATA_DIR / "example_wb.xlsx",
        cells_to_change = {"B1": 33, "B2": 44},
        cells_to_retrieve=['B4', 'B5'],
        sheet_idx=0,
    )
    assert results['B4'] == 22
    assert results['B5'] == 11

    results2 = xl.execute_workbook(
        TEST_DATA_DIR / "example_wb.xlsx",
        cells_to_change = {"B1": 33, "B2": 66},
        cells_to_retrieve={'B4': "label1", 'B5': "label2"},
        sheet_idx=1,
        new_filepath=TEST_DATA_DIR / "stored_results.xlsx"
    )
    assert results2['label1'] == 44
    assert math.isclose(results2['label2'], 39.6)

    import xlwings as xw
    with xw.Book(TEST_DATA_DIR / "stored_results.xlsx") as wb:
        ws = wb.sheets[1]
        assert ws["B1"].value == 33
        assert ws["B2"].value == 66

def test_excel_runner():
    dcr2 = xl.create_condition_check(2, "ge")
    results = xl.excel_runner(
        TEST_DATA_DIR / "example_wb.xlsx",
        static_inputs={"B1": [10, 20], "Labels": ["C01", "C02"]},
        dynamic_inputs={
            "OptA": {"B2": 22},
            "OptB": {"B2": 33},
            "OptC": {"B2": 55},
        },
        save_conditions={"B6": dcr2},
        static_identifier_keys=["Labels"],
        result_labels={"B6": "meaningful_value"},
        save_dir=TEST_DATA_DIR / "design"
    )
    assert results['C01']['successful_key'] is None # No match found
    assert results['C02']['successful_key'] == "OptA" # Option A worked passed the criteria