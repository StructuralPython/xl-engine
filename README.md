# xl-engine
 Use your existing Excel workbooks like they were Python programs: A utility built on XLWings to automate the input, execution, and data retrieval of your existing Excel Workbooks. 


## Installation

`pip install xl-engine`

## Basic Usage

### `execute_workbook`

Use `execute_workbook` when you want to execute a workbook once with known, provided parameters.

```python
import xl_engine as xl

results = xl.execute_workbook(
    TEST_DATA_DIR / "example_wb.xlsx",
    cells_to_change = {"B1": 33, "B2": 66},
    cells_to_retrieve=['B4', 'B5'], # This can either be a list or a dict; see next example
    sheet_idx=1,
    new_filepath=TEST_DATA_DIR / "stored_results.xlsx"
)

# results
# {'B4': 22.0, 'B5': 11.0}
```

You can also use a dictionary in `cells_to_retrieve` to meaningfully label the results:

```python
results2 = xl.execute_workbook(
    TEST_DATA_DIR / "example_wb.xlsx",
    cells_to_change = {"B1": 33, "B2": 66},
    cells_to_retrieve={'B4': "label1", 'B5': "label2"}, # Now a dict
    sheet_idx=1,
    new_filepath=TEST_DATA_DIR / "stored_results.xlsx"
)

# results2
# {'label1': 44.0, 'label2': 39.599999999999994}
```

### `excel_runner`

Use `excel_runner` when you want to execute a workbook multiple times (static inputs) with multiple options (dynamic inputs).

```python
import xl_engine as xl

# Creates a callable with the following function signature: callable(x: float | int) -> bool:
# When a value is a passed to dcr2, i.e. dcr2(some_value), it will return True if
# the value is greater-than-or-equal-to 2
dcr2 = xl.create_condition_check(2, "ge")

# static_inputs will be used to populate each workbook.
# Within each set of static inputs, a second sub-iteration will occur
# where the dynamic inputs will be input allowing you to test the
# results for different "options" of dynamic inputs.
# If the save conditions pass (return True for all keys), then a copy of the
# workbook will be saved to disk with the successful inputs recorded.
# A dictionary of all calculated results will be returned.
results = xl.excel_runner(
        TEST_DATA_DIR / "example_wb.xlsx",
        static_inputs={"B1": [10, 20], "Labels": ["C01", "C02"]},
        dynamic_inputs={
            "OptA": {"B2": 22},
            "OptB": {"B2": 33},
            "OptC": {"B2": 55},
        },
        success_conditions={"B6": dcr2},
        static_identifier_keys=["Labels"],
        result_labels={"B6": "meaningful_value"},
        save_dir=TEST_DATA_DIR / "design"
)
```

Return value (for the "example_wb.xlsx" file included in the tests directory):

```python
{
    'C01': {
        'successful_key': None, # None of the options were successful for this case
        'OptA': {'meaningful_value': 1.2121212121212122}, 
        'OptB': {'meaningful_value': 0.8080808080808081}, 
        'OptC': {'meaningful_value': 0.48484848484848486}
    }, 
    'C02': {
        'successful_key': 'OptA', 
        'meaningful_value': 2.4242424242424243
    }
}
```