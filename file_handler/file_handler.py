import os
import shutil
import tempfile
import zipfile


class FileHandler:
    def __init__(self, filepath, converter, config=None):
        self._filepath = filepath
        self._converter = converter

        self._tmp_dir = tempfile.mkdtemp()
        with open(filepath, 'rb') as f:
            zip_file = zipfile.ZipFile(f)
            zip_file.extractall(self._tmp_dir)
            self._filenames = zip_file.namelist()

    def _export_file(self):
        with zipfile.ZipFile(self._filepath, 'w') as docx:
            for filename in self._filenames:
                docx.write(os.path.join(self._tmp_dir, filename), filename)

    def process_file(self):
        pass

    def __del__(self):
        try:
            shutil.rmtree(self._tmp_dir)
        except Exception:
            pass
