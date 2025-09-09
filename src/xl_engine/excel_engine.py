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
    demand_input_cell_arrays: dict[str, list],
    design_inputs: dict[str, dict[str, float]],
    result_cells: list[str],
    save_conditions: dict[str, callable],
    identifier_keys: Optional[list[str]] = None,
    save_dir: Optional[str] = None,
    sheet_idx: int = 0,
) -> None:
    """
    Doc strign
    """

    demand_cell_ids = list(demand_input_cell_arrays.keys())
    iterations = len(demand_input_cell_arrays[demand_cell_ids[0]])

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
    with Live(panel) as live:
        for iteration in range(iterations):
            demand_cells_to_change = {
                cell_id: demand_input_cell_arrays[cell_id][iteration] 
                for cell_id in demand_cell_ids
                if valid_excel_reference(cell_id)
            }
            identifier_values = {
                cell_id: str(demand_input_cell_arrays[cell_id][iteration])
                for cell_id in demand_cell_ids
                if not valid_excel_reference(cell_id)
            }
            variations_task = variations_progress.add_task("Sheet variations", total=len(design_inputs.items()))
            variations_progress.reset(variations_task)
            for design_tag, design_cells_to_change in design_inputs.items():
                cells_to_change = demand_cells_to_change | design_cells_to_change
                calculated_results = execute_workbook(
                    xlsx_filepath, 
                    cells_to_change=cells_to_change,
                    cells_to_retrieve=result_cells,
                    sheet_idx=sheet_idx
                )
            
                save_condition_acc = []
                for result_cell_id, save_condition in save_conditions.items():
                    calculated_result = calculated_results[result_cell_id]
                    save_condition_acc.append(save_condition(calculated_result))
                variations_progress.update(variations_task, advance=1)
                
                if all(save_condition_acc):
                    filepath = pathlib.Path(xlsx_filepath)
                    name = filepath.stem
                    suffix = filepath.suffix
                    if identifier_values:
                        identifiers = "-".join([identifier_values[id_key] for id_key in identifier_values])
                    else:
                        identifiers = f"{iteration}"
                    
                    new_filename = f"{name}-{identifiers}-{design_tag}{suffix}"
                    save_dir_path = pathlib.Path(save_dir)
                    if not save_dir_path.exists():
                        save_dir_path.mkdir(parents=True)
                    calculated_results = execute_workbook(
                        xlsx_filepath, 
                        cells_to_change=cells_to_change,
                        cells_to_retrieve=result_cells,
                        new_filepath=f"{str(save_dir)}/{new_filename}",
                        sheet_idx=sheet_idx,
                    )
                    variations_progress.remove_task(variations_task)
                    break
            else:
                variations_progress.remove_task(variations_task)
                progress_group.renderables.append(Text(f"Variation: {iteration} did not meet the criteria"))
            main_progress.update(main_task, advance=1)


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
    """
    xlsx_filepath = pathlib.Path(xlsx_filepath)
    if not xlsx_filepath.exists():
        raise FileNotFoundError(f"Please check your file location since this does not exist: {xlsx_filepath}")
    with xw.App(visible=False) as app:
        wb = xw.Book(xlsx_filepath)
        ws = wb.sheets[sheet_idx]
        for cell_name, new_value in cells_to_change.items():
            try:
                ws[cell_name].value = new_value
            except:
                raise ValueError(f"Invalid input cell name: {cell_name}. Perhaps you made a typo?")
    
        calculated_values = {} # Add afterwards
        for cell_to_retrieve in cells_to_retrieve:
            try:
                retrieved_value = ws[cell_to_retrieve].value
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