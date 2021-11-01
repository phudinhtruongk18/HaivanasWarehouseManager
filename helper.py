import os


def create_export_fordel(dl_path):
    if not os.path.exists(dl_path):
            os.mkdir(dl_path)