import xlsxwriter
import ruamel.yaml as yaml
from pathlib import Path
import importlib_resources
from collections import OrderedDict


class ExcelWriter:
    def __init__(self, data, config=None):
        self.config = None

        if config is not None:
            if isinstance(config, str):
                config_yaml_path = Path(config)
                if config_yaml_path.exists() and config_yaml_path.suffix in ('.yml', '.yaml'):
                    self.config = yaml.safe_load(config_yaml_path.read_text())
            elif isinstance(config, (dict, OrderedDict)):
                self.config = config

        if self.config is None:
            self.config = yaml.safe_load(importlib_resources.read_text('pyexcel_xlsxwx', 'default.yaml'))

        self.data = data

    def save(self, out_file):
        wb = xlsxwriter.Workbook(out_file, self.config.get('workbook', dict()))

        for sheet_name in self.data.keys():
            ws = wb.add_worksheet(sheet_name)

        self.set_worksheet_formatting(wb)
        self.set_formatting(wb)

        for sheet_name, sheet_matrix in self.data.items():
            ws = wb.get_worksheet_by_name(sheet_name)

            for row_i, row in enumerate(sheet_matrix):
                ws.write_row(row_i, 0, row)

        wb.close()

    def set_formatting(self, wb):
        format_config = self.config.get('format', dict())

        default_format = format_config.pop('_default', None)
        if default_format is not None:
            default_format = wb.add_format(default_format)

            for sheet_name, sheet_matrix in self.data.items():
                ws = wb.get_worksheet_by_name(sheet_name)

                for row_i, _ in enumerate(sheet_matrix):
                    ws.set_row(row_i, None, default_format)

        for sheet_name, sheet_format in format_config.items():
            ws = wb.get_worksheet_by_name(sheet_name)

            default_format = sheet_format.pop('_default', None)
            if default_format is not None:
                default_format = wb.add_format(default_format)

                for row_i, _ in enumerate(self.data[sheet_name]):
                    ws.set_row(row_i, None, default_format)

            cell_format = dict()
            for position, formatting in sheet_format.items():
                if position.isdigit() or isinstance(position, int):
                    row_format = wb.add_format(formatting)
                    ws.set_row(int(position) - 1, None, row_format)
                elif position.isalpha():
                    col_format = wb.add_format(formatting)
                    ws.set_column('{0}:{0}'.format(position), None, col_format)
                else:
                    cell_format[position] = wb.add_format(formatting)

            for position, formatting in cell_format.items():
                ws.write_blank(position, None, formatting)

    def set_worksheet_formatting(self, wb):
        worksheet_config = self.config.get('worksheet', dict())

        default_format = worksheet_config.pop('_default', None)
        if default_format is not None:
            for sheet_name in self.data.keys():
                ws = wb.get_worksheet_by_name(sheet_name)

                freeze_panes = default_format.get('freeze_panes', None)
                if freeze_panes is not None:
                    ws.freeze_panes(freeze_panes)

                column_width = default_format.get('column_width', None)
                if column_width is not None:
                    if isinstance(column_width, list):
                        for i, width in enumerate(column_width):
                            ws.set_column(i, i, width)
                    else:
                        ws.set_column(0, len(self.data[sheet_name][0]) - 1, column_width)

        for sheet_name, formatting in worksheet_config.items():
            ws = wb.get_worksheet_by_name(sheet_name)

            freeze_panes = formatting.get('freeze_panes', None)
            if freeze_panes is not None:
                ws.freeze_panes(freeze_panes)

            column_width = formatting.get('column_width', None)
            if column_width is not None:
                if isinstance(column_width, list):
                    for i, width in enumerate(column_width):
                        ws.set_column(i, i, width)
                else:
                    ws.set_column(0, len(self.data[sheet_name][0]) - 1, column_width)