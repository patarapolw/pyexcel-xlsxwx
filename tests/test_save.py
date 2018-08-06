import pytest
from pathlib import Path
import pyexcel

import pyexcel_xlsxwx


@pytest.mark.parametrize('in_file', [
    'test.xlsx'
])
@pytest.mark.parametrize('config', [
    None,
    'config1.yaml'
])
def test_save(in_file, config,
              request):
    config_path = None
    if config is not None:
        config_path = Path('tests/input').joinpath(config)
        assert config_path.exists()

        config_path = str(config_path)

    data = pyexcel.get_book_dict(file_name=str(Path('tests/input').joinpath(in_file)))
    pyexcel_xlsxwx.save_data(str(Path('tests/output').joinpath(request.node.name).with_suffix('.xlsx')),
                             data, config=config_path)
