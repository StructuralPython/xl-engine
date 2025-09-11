from typing import Optional
import operator
import pathlib
import re
import xlwings as xw
from rich.text import Text
from rich.progress import Progress, TextColumn, SpinnerColumn, BarColumn, MofNCompleteColumn, TaskProgressColumn, TimeRemainingColumn
from rich.panel import Panel
from rich.padding import Padding
from rich.console import Group
from rich.live import Live


def excel_runner(
    xlsx_filepath,
    static_inputs: dict[str, list[float | int | str]] | list[dict[str, float | int | str]],
    dynamic_inputs: dict[str, dict[str, float]],
    success_conditions: dict[str, callable],
    static_identifier_keys: Optional[list[str]] = None,
    result_labels: Optional[dict[str, str]] = None,
    static_input_maps: Optional[dict[str, str]] = None,
    save_dir: Optional[str] = None,
    sheet_idx: int = 0,
) -> None:
    """
    Executes the Excel workbook at xlsx_filepath with each static_input and then for
    each dynamic_input. The 'save_conditions' dictionary is a keyed by cell references
    representing result cells. The values for the save_conditions are unary callables
    (single argument) that return a bool. If all callables in 'save_conditions' return
    True for that iteration, then the workbook with that iteration's inputs is saved
    to disk. 

    'static_inputs': a dictionary where each key is a cell reference and the values are a
        list of values to place in the cell OR a list of dictionaries, where each key is a
        cell reference and the value is the value to be placed.
        Each index in the list (whether the list is the inner or outer collection type)
        represents the data used for one iteration. For keys containing cell references,
        the cell reference is used as a static input value to the workbook. For keys
        that are not cell references, their values are accessible by 'static_identifier_keys',
        which will be used for creating the unique filename in the event that the
        'save_conditions' are satisfied. If the key is not in 'static_identifier_keys',
        it will be assumed to be a cell reference.
    'dynamic_inputs': a dictionary of dictionaries. The outer keys represent the unique
        label to describe the iteration, e.g. the name of the design element.
        The values are dictionaries keyed by cell references with single values which will
        populate the workbook for every static iteration.
    'success_conditions': a dictionary keyed by cell references whose values are unary callables
        which return a bool when passed the value retrieved from the workbook at the cell 
        reference during each iteration. If all callables in the 'success_conditions' dict
        return True, then that iteration of the workbook will be saved to disk. Use the
        create_condition_check() function in this module to quickly create such callables.
    'static_identifier_keys': The keys in 'static_inputs', which are not cell references,
        which will be used to create unique filenames whenever the save conditions are all
        True. e.g. if there is a key in 'static_inputs' called 'Label', then passing a list
        of ['Label'] as the static_identifier_keys will ensure that the data under the 'Label'
        key will be used as part of the unique filename.
    'result_labels': A mapping of result cell references to what those cell references mean
        for a human. e.g. if the result cell B6 referred to a "shear utilization ratio" then
        the result_label dict might look like this: {"B6": "shear utilization ratio"}
        The result label will be used in the returned results. If None, then the cell references
        will be used instead.
    'static_input_maps': A mapping of keys in 'static_inputs' to cell references. This is useful
        to provide when your data is keyed by some other process and you do not want to manually
        re-map your data to be keyed by cell referdences. By providing 'static_input_maps', this
        excel_runner will do that for you.
    'save_dir': The directory to store saved workbooks
    'sheet_idx': The sheet to modify within the workbook.
    """
    static_inputs = format_static_inputs(static_inputs)
    iterations = len(static_inputs)
    if static_input_maps is None:
        static_input_maps = dict()

    main_progress = Progress(
    TextColumn("{task.description}"),
    BarColumn(),
    MofNCompleteColumn(),
    TaskProgressColumn(),
    TimeRemainingColumn(),
    expand=True,
    )
    variations_progress = Progress(
        TextColumn("{task.description}"),
        MofNCompleteColumn(),
        BarColumn(),
    )
    indented_variations = Padding(variations_progress, (0, 0, 0, 4))
    progress_group = Group(main_progress, indented_variations)
    panel = Panel(progress_group)

    main_task = main_progress.add_task("Primary Iterations", total=iterations)
    dynamic_results = {}
    with Live(panel) as live:
        for iteration in range(iterations):
            static_data = static_inputs[iteration]
            demand_cells_to_change = {
                static_input_maps.get(k, k): v
                for k, v in static_data.items()
                if k not in static_identifier_keys
            }
            identifier_values = {
                k: str(v)
                for k, v in static_data.items()
                if k in static_identifier_keys
            }
            if identifier_values:
                identifiers = "-".join(identifier_values.values())
            else:
                identifiers = f"{iteration}"
            variations_task = variations_progress.add_task("Sheet variations", total=len(dynamic_inputs.items()))
            variations_progress.reset(variations_task)
            failed_results = {}
            for design_tag, design_cells_to_change in dynamic_inputs.items():
                cells_to_change = demand_cells_to_change | design_cells_to_change
                calculated_results = execute_workbook(
                    xlsx_filepath, 
                    cells_to_change=cells_to_change,
                    cells_to_retrieve=list(success_conditions.keys()),
                    sheet_idx=sheet_idx
                )
                if isinstance(result_labels, dict):
                    labeled_results = {result_labels[k]: v for k, v in calculated_results.items()}
                else:
                    labeled_results = calculated_results

                failed_results.update({design_tag: labeled_results})

                save_condition_acc = []
                for result_cell_id, save_condition in success_conditions.items():
                    calculated_result = calculated_results[result_cell_id]
                    save_condition_acc.append(save_condition(calculated_result))
                variations_progress.update(variations_task, advance=1)
                
                if all(save_condition_acc):
                    filepath = pathlib.Path(xlsx_filepath)
                    name = filepath.stem
                    suffix = filepath.suffix
                    
                    new_filename = f"{name}-{identifiers}-{design_tag}{suffix}"
                    save_dir_path = pathlib.Path(save_dir)
                    if not save_dir_path.exists():
                        save_dir_path.mkdir(parents=True)
                    _ = execute_workbook(
                        xlsx_filepath, 
                        cells_to_change=cells_to_change,
                        cells_to_retrieve=list(success_conditions.keys()),
                        new_filepath=f"{str(save_dir)}/{new_filename}",
                        sheet_idx=sheet_idx,
                    )
                    dynamic_results.update({identifiers: {"successful_key": design_tag} | labeled_results})
                    variations_progress.remove_task(variations_task)
                    break
            else:
                variations_progress.remove_task(variations_task)
                dynamic_results.update({identifiers: {"successful_key": None} | failed_results})
                progress_group.renderables.append(Text(f"Variation: {iteration} did not meet the criteria"))
            main_progress.update(main_task, advance=1)
    return dynamic_results


def execute_workbook(
        xlsx_filepath: str | pathlib.Path, 
        cells_to_change: dict[str, str | float | int], 
        cells_to_retrieve: list[str], 
        sheet_idx=0,
        new_filepath: Optional[str | pathlib.Path] = None, 
) -> dict:
    """
    Executes the Excel workbook located at 'xlsx_filepath' after it has been populated
    with the data in 'cells_to_change'. Returns the values of 'cells_to_retrieve' as a
    dictionary of values retrieved from the executed notebook.

    'xlsx_filepath': A path to an existing Excel workbook. Can be relative or absolute
        path in either str form or pathlib.Path. If you are using backslashes as part 
        of a filepath str on Windows, make sure they are escaped.
    'cells_to_change': A dictionary where the keys are Excel cell names (e.g. "E45")
        and the values are the values that should be set for each key.
    'cells_to_retrieve': Either a list or dict. If list, represents a list of str 
        representing Excel cell names that should be retrieved after computation 
        (e.g. ['C1', 'E5']).
        If a dict, the keys are the cell references and the values are what the 
        cell references represent. The values will be used as the keys in the 
        returned dictionary. (e.g. {"C1": "Date", "E5": "Critical value"})
    'sheet_idx': The sheet in the workbook 
    'new_filepath': If not None, a copy of the altered workbook will be saved at this
        locations. Can be a str or pathlib.Path. Directories on
        the path must already exist because this function will not create them if
        they do not.

    ### Example:

        dcr2 = xl.create_condition_check(2, "ge")
        results = xl.excel_runner(
            "example_wb.xlsx",
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
    """
    xlsx_filepath = pathlib.Path(xlsx_filepath)
    if not xlsx_filepath.exists():
        raise FileNotFoundError(f"Please check your file location since this does not exist: {xlsx_filepath}")
    with xw.App(visible=False) as app:
        wb = xw.Book(xlsx_filepath)
        ws = wb.sheets[sheet_idx]
        for cell_name, new_value in cells_to_change.items():
            try:
                ws.range(cell_name).value = new_value
            except:
                raise ValueError(f"Invalid input cell name: {cell_name}. Perhaps you made a typo?")

        calculated_values = {} # Add afterwards
        for cell_to_retrieve in cells_to_retrieve:
            try:
                retrieved_value = ws.range(cell_to_retrieve).value
            except:
                raise ValueError(f"Invalid retrieval cell name: {cell_to_retrieve}. Perhaps you made a typo?")
            label = cell_to_retrieve
            if isinstance(cells_to_retrieve, dict):
                label = cells_to_retrieve.get(cell_to_retrieve, cell_to_retrieve)
            calculated_values.update({label: retrieved_value})
    
        if new_filepath:
            new_filepath = pathlib.Path(new_filepath)
            if not new_filepath.parent.exists():
                raise FileNotFoundError(f"The parent directory does not exist: {new_filepath.parent}")
            try:
                wb.save(new_filepath)
            except Exception as e:
                print(e)
                raise RuntimeError(
                    "An error occured with the Excel interface during saving. Possible causes include:\n"
                    "- You do not have permissions to save to the chosen location.\n"
                    "- Your hard-drive is full.\n"
                )
        wb.close()
    return calculated_values


def create_condition_check(check_against_value: float, op: str) -> callable:
    """
    Returns a function with a single numerical input parameter.
    The function returns a boolean corresponding to whether the 
    single numerical argument passed to it meets the condition
    encoded in the function.

    'check_against_value' the value that will be encoded in the function
        to check against.
    'op': str, one of {"ge", "le", "gt", "lt", "eq", "ne"}
        - "ge" Greater-than-or-equal-to
        - "le" Less-than-or-equal-to
        - "gt" Greater-than
        - "lt" Less-than
        - "eq" Equal-to
        - "ne" Not-equal-to
    """
    operators = {
        "ge": operator.ge,
        "le": operator.le,
        "gt": operator.gt,
        "lt": operator.lt,
        "eq": operator.eq,
        "ne": operator.ne,
    }
    def checker(test_value):
        return operators[op.lower()](test_value, check_against_value)

    return checker
    

def valid_excel_reference(cell: str) -> bool:
    """
    Returns True if 'cell' is a value that represents a valid
    MS Excel reference, e.g. "B4", "AAC93290"
    """
    pattern = re.compile("^[A-Z]{1,3}[0-9]+$")
    match = pattern.match(cell)
    if match is None:
        return False
    else:
        return True
    

def format_static_inputs(
        static_inputs: dict[str, list[float | int | str]] | list[dict[str, float | int | str]]
) -> list[dict[str, float | int | str]]:
    """
    Transforms a dictionary of str keys and list values to a list of dictionaries.

    All sub-lists must be the same size.
    """
    if isinstance(static_inputs, list) and isinstance(static_inputs[0], dict):
        return static_inputs
    else:
        column_data = [list_data for list_data in static_inputs.values()]
        row_data = zip(*column_data)
        outer_acc = []
        for row in row_data:
            inner_acc = {}
            for idx, key in enumerate(static_inputs.keys()):
                inner_acc.update({key: row[idx]})
            outer_acc.append(inner_acc)
    return outer_acc

