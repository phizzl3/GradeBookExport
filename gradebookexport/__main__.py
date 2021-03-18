from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder

import dropfile
from xlclass import Xlsx


class Grades(Xlsx):

    def __init__(self):
        super().__init__()

    def set_column_width(self, column_width=15):
        """
        Adjust all columns to the same specified width.
        """
        dim_holder = DimensionHolder(worksheet=xl.ws)
        for col in range(xl.ws.min_column, xl.ws.max_column + 1):
            dim_holder[get_column_letter(col)] = ColumnDimension(
                xl.ws, min=col, max=col, width=column_width)
        xl.ws.column_dimensions = dim_holder

    def remove_ungraded_columns(self):
        """
        Remove any columns that contain "Ungraded" content...multi-pass
        because I'm lazy.
        """
        for i in range(5):
            for col in xl.ws.iter_cols():
                for cell in col:
                    if cell.value:
                        if 'Ungraded' in cell.value:
                            xl.ws.delete_cols(cell.column, 1)


if __name__ == '__main__':
    _csv = dropfile.get()
    xl = Grades()
    xl.copy_csv_data(_csv)
    xl.set_column_width()
    xl.remove_ungraded_columns()
    xl.save(f'{_csv.parent}/{_csv.stem}.xlsx')
