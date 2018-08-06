# pyexcel-xlsxwx

[![Build Status](https://travis-ci.org/patarapolw/pyexcel-xlsxwx.svg?branch=master)](https://travis-ci.org/patarapolw/pyexcel-xlsxwx)
[![PyPI version shields.io](https://img.shields.io/pypi/v/pyexcel_xlsxwx.svg)](https://pypi.python.org/pypi/pyexcel_xlsxwx/)
[![PyPI license](https://img.shields.io/pypi/l/pyexcel_xlsxwx.svg)](https://pypi.python.org/pypi/pyexcel_xlsxwx/)
[![PyPI pyversions](https://img.shields.io/pypi/pyversions/pyexcel_xlsxwx.svg)](https://pypi.python.org/pypi/pyexcel_xlsxwx/)

Save pyexcel data with XlsxWriter, while retaining good formatting.

## Features

- Allow setting column widths and word wrap.
- A package for reading data is not included, please see [`pyexcel`'s plugins here](https://github.com/pyexcel/pyexcel#available-plugins).

## Installation

```commandline
$ pip install pyexcel-xlsxwx
```
Note that `pyexcel` is not a dependency.

## Usage

```python
>>> import pyexcel_xlsxwx
>>> data = OrderedDict() # from collections import OrderedDict
>>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
>>> data.update({"Sheet 2": [["row 1", "row 2", "row 3"]]})
>>> pyexcel_xlsxwx.save_data("your_file.xlsx", data)
```

You can also define a custom config via:
```python
>>> pyexcel_xlsxwx.save_data("your_file.xlsx", data, config=config)
```
Where config can be dictionary or path to YAML file.

The default YAML config is:

```yaml
workbook:
  constant_memory: true
  strings_to_numbers: false
  strings_to_formulas: false
  strings_to_urls: true
worksheet:
  _default:
    freeze_panes: A2
#    column_width: 30
    smart_fit: true
    max_column_width: 30
format:
  _default:
    valign: top
    text_wrap: true
```
`column_width` can also accept a list of numbers.

For example, to cancel out `freeze_panes`, try:

```python
>>> pyexcel_xlsxwx.save_data("your_file.xlsx", data, config={'worksheet': {'_default': {'freeze_panes': None}}})
```

The settings will merge (thanks to https://stackoverflow.com/questions/20656135/python-deep-merge-dictionary-data), so that the other formattings won't be lost.

## Associated projects

- [pyexcel-export](https://github.com/patarapolw/pyexcel-export) - operates using OpenPyXL, which seeming has bad word wrap support. However, the formatting can be well preserved.
