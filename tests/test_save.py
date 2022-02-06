import pytest
from pathlib import Path
import pyexcel

import pyexcel_xlsxwx


@pytest.mark.parametrize("in_file", ["test.xlsx"])
@pytest.mark.parametrize(
    "config",
    [None, "config1.yaml", {"worksheet": {"_default": {"freeze_panes": None}}}],
)
def test_save(in_file, config, request):
    if isinstance(config, str):
        config = Path("tests/input").joinpath(config)
        assert config.exists()

        config = str(config)

    data = pyexcel.get_book_dict(file_name=str(Path("tests/input").joinpath(in_file)))
    pyexcel_xlsxwx.save_data(
        str(Path("tests/output").joinpath(request.node.name).with_suffix(".xlsx")),
        data,
        config=config,
    )
