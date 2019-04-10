import openpyxl
from openpyxl.utils import get_column_letter
import ruamel.yaml as yaml
from pathlib import Path
# from copy import copy


class ExcelReader:
    def __init__(self, filename):
        self.wb = openpyxl.load_workbook(filename)

    @staticmethod
    def _to_styles(wb, skip_cell_formatting=True):
        """

        :param openpyxl.workbook.workbook.Workbook wb:
        :param skip_cell_formatting:
        :return:
        """
        styles = {
            'worksheet': dict()
        }

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            styles['worksheet'][sheet_name] = {
                'freeze_panes': ws.freeze_panes
            }

            for i, col in enumerate(ws.iter_cols()):
                styles['worksheet'][sheet_name].setdefault('column_width', []) \
                    .append(ws.column_dimensions[get_column_letter(i + 1)].width)

            for i, row in enumerate(ws):
                styles['worksheet'][sheet_name].setdefault('row_height', []) \
                    .append(ws.row_dimensions[i + 1].height)

                # if not skip_cell_formatting:
                #     for cell in row:
                #         font = cell.font
                #         font_color = font.color
                #
                #         alignment = cell.alignment
                #         fill = cell.fill
                #         border = cell.border
                #
                #         cell_format = {
                #             'font': {
                #                 'font_name': font.name,
                #                 'font_size': font.sz,
                #                 'font_color': copy(font_color.rgb) if font_color is not None else None,
                #                 'bold': font.b,
                #                 'italic': font.i,
                #                 'underline': font.u,
                #                 'font_strikeout': font.strike,
                #                 # 'font_script':
                #             },
                #             'number': {
                #                 'num_format': cell.number_format
                #             },
                #             # 'Protection':
                #             'alignment': {
                #                 'align': alignment.horizontal,
                #                 'valign': alignment.vertical,
                #                 'rotation': alignment.textRotation,
                #                 'text_wrap': alignment.wrapText,
                #                 'reading_order': alignment.readingOrder,
                #                 'text_justlast': alignment.justifyLastLine,
                #                 # 'center_across':,
                #                 'indent': alignment.relativeIndent,
                #                 'shrink': alignment.shrinkToFit
                #             },
                #             'pattern': {
                #                 'pattern': fill.patternType,
                #                 'bg_color': fill.bgColor.rgb if fill.bgColor is not None else None,
                #                 'fg_color': fill.fgColor.rgb if fill.bgColor is not None else None
                #             },
                #             'border': {
                #                 # 'border':,
                #                 # 'border_color':
                #                 'bottom': border.bottom.style,
                #                 'bottom_color': border.bottom.color.rgb if border.bottom.color is not None else None,
                #                 'top': border.top.style,
                #                 'top_color': border.top.color.rgb if border.bottom.color is not None else None,
                #                 'left': border.left.style,
                #                 'left_color': border.left.color.rgb if border.bottom.color is not None else None,
                #                 'right': border.right.style,
                #                 'right_color': border.right.color.rgb if border.bottom.color is not None else None
                #             }
                #         }
                #
                #         styles.setdefault('format', dict())\
                #             .setdefault(sheet_name, dict())\
                #             [cell.coordinate] = cell_format

            if all(height is None for height in styles['worksheet'][sheet_name]['row_height']):
                styles['worksheet'][sheet_name].pop('row_height')

        return styles

    def get_styles(self, out_file=None, skip_cell_formatting=True):
        formatting = self._to_styles(self.wb, skip_cell_formatting)

        if out_file is not None:
            return Path(out_file).write_text(yaml.safe_dump(formatting))
        else:
            return formatting
