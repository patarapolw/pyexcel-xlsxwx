from .app import ExcelWriter


def save_data(out_file, data, config=None):
    ExcelWriter(data, config=config).save(out_file)
