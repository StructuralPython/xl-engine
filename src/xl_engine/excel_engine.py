from typing import Optional
import operator
import pathlib
import xlwings as xw


def excel_runner(
    xlsx_filepath,
    demand_input_cell_arrays: dict[str, list],
    identifier_cell_arrays: dict[str, list],
    design_inputs: dict[str, dict[str, float]],
    result_cells: list[str],
    save_conditions: dict[str, callable],
    save_dir: Optional[str] = None,
    sheet_idx: int = 0
) -> None:
    demand_cell_ids = list(demand_input_cell_arrays.keys())
    iterations = len(demand_input_cell_arrays[demand_cell_ids[0]])
    for iteration in range(iterations):
        demand_cells_to_change = {cell_id: demand_input_cell_arrays[cell_id][iteration] for cell_id in demand_cell_ids}
        for design_tag, design_cells_to_change in design_inputs.items():
            cells_to_change = demand_cells_to_change | design_cells_to_change
            calculated_results = excel_engine(
                xlsx_filepath, 
                cells_to_change=cells_to_change,
                cells_to_retrieve=result_cells,
                sheet_idx=sheet_idx
            )
        
            save_condition_acc = []
            for idx, result_cell_id in enumerate(result_cells):
                calculated_result = calculated_results[idx]
                save_condition_acc.append(save_conditions[result_cell_id](calculated_result))
            
            if all(save_condition_acc):
                filepath = pathlib.Path(xlsx_filepath)
                name = filepath.stem
                suffix = filepath.suffix
                demand_ids = "-".join([id_array[iteration] for id_array in identifier_cell_arrays])
                
                new_filename = f"{name}-{demand_ids}-{design_tag}.{suffix}"
                save_dir_path = pathlib.Path(save_dir)
                if not save_dir_path.exists():
                    save_dir_path.mkdir(parents=True)
                calculated_results = excel_engine(
                    xlsx_filepath, 
                    cells_to_change=cells_to_change,
                    cells_to_retrieve=result_cells,
                    sheet_idx=sheet_idx,
                    new_filename=f"{str(save_dir)}/{new_filename}"
                )
                break


def excel_engine(xlsx_file_name, cells_to_change, cells_to_retrieve, new_file_name="", sheet_idx=0):
    """
    Returns a list of the updated cell values for the cell ids in 'cells_to_retrieve'
    """
    with xw.App(visible=False) as app:
        wb = xw.Book(xlsx_file_name)
        ws = wb.sheets[sheet_idx]
        for cell_name, new_value in cells_to_change.items():
            ws[cell_name].value = new_value
    
        calculated_values = [] # Add afterwards
        for cell_to_retrieve in cells_to_retrieve:
            retrieved_value = ws[cell_to_retrieve].value
            calculated_values.append(retrieved_value)
    
        if new_file_name:
            wb.save(new_file_name)
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
        return operators[op](test_value, check_against_value)

    return checker
    