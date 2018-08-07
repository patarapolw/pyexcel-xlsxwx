import xlsxwriter
import ruamel.yaml as yaml
from pathlib import Path
import importlib_resources
from collections import OrderedDict


class ExcelWriter:
    def __init__(self, data, config=None):
        self.config = dict()

        if config is not None:
            if isinstance(config, str):
                config_yaml_path = Path(config)
                if config_yaml_path.exists() and config_yaml_path.suffix in ('.yml', '.yaml'):
                    self.config = yaml.safe_load(config_yaml_path.read_text())
            elif isinstance(config, (dict, OrderedDict)):
                self.config = config

        self.config = deep_merge_dict(self.config,
                                      yaml.safe_load(importlib_resources.read_text('pyexcel_xlsxwx', 'default.yaml')))

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
                self._set_worksheet_formatting(wb, sheet_name, default_format)

        for sheet_name, formatting in worksheet_config.items():
            self._set_worksheet_formatting(wb, sheet_name, formatting)

    def _set_worksheet_formatting(self, wb, sheet_name, formatting):
        ws = wb.get_worksheet_by_name(sheet_name)

        freeze_panes = formatting.get('freeze_panes', None)
        if freeze_panes is not None:
            ws.freeze_panes(freeze_panes)

        smart_fit = formatting.get('smart_fit', False)
        max_column_width = formatting.get('max_column_width', 30)
        if smart_fit:
            for col_i, _ in enumerate(self.data[sheet_name][0]):
                col_width = max([len(str(row[col_i])) if row[col_i] is not None else 0
                                 for row in self.data[sheet_name]]) + 2
                if col_width > max_column_width:
                    col_width = max_column_width

                ws.set_column(col_i, col_i, col_width)

        column_width = formatting.get('column_width', None)
        if column_width is not None:
            if isinstance(column_width, list):
                for i, width in enumerate(column_width):
                    ws.set_column(i, i, width)
            elif isinstance(column_width, (dict, OrderedDict)):
                for key, width in column_width.keys():
                    ws.set_column('{0}:{0}'.format(key), width)
            else:
                ws.set_column(0, len(self.data[sheet_name][0]) - 1, column_width)

        row_height = formatting.get('row_height', None)
        if row_height is not None:
            if isinstance(row_height, list):
                for i, height in enumerate(row_height):
                    ws.set_row(i, i, height)
            elif isinstance(row_height, (dict, OrderedDict)):
                for key, height in row_height.items():
                    ws.set_row(int(key), int(key), height)
            else:
                ws.set_row(0, len(self.data[sheet_name]) - 1, row_height)


def deep_merge_dict(source, destination):
    """
    run me with nosetests --with-doctest file.py

    >>> a = { 'first' : { 'all_rows' : { 'pass' : 'dog', 'number' : '1' } } }
    >>> b = { 'first' : { 'all_rows' : { 'fail' : 'cat', 'number' : '5' } } }
    >>> deep_merge_dict(b, a) == { 'first' : { 'all_rows' : { 'pass' : 'dog', 'fail' : 'cat', 'number' : '5' } } }
    True
    """
    for key, value in source.items():
        if isinstance(value, (dict, OrderedDict)):
            # get node or create one
            node = destination.setdefault(key, {})
            deep_merge_dict(value, node)
        else:
            destination[key] = value

    return destination
