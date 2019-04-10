from .app import ExcelReader


def get_styles(in_file, out_file=None, skip_cell_formatting=True):
    return ExcelReader(in_file).get_styles(out_file=out_file, skip_cell_formatting=skip_cell_formatting)
