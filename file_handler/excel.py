import os
import shutil
import subprocess
import tempfile
from lxml import etree

import constants.constants as constants
import constants.errors as errors
import utils.helper as helper
from constants.charset import Encoding
from file_handler.file_handler import FileHandler


class XlsxFileHandler(FileHandler):
    def _convert_encoding(self, xml_filepath):
        content_tree = etree.parse(xml_filepath)

        for r in content_tree.xpath(".//*[local-name() = 't']"):
            r.text, _ = self._converter.convert(r.text)

        content_tree.write(xml_filepath)

    def _convert_style_file(self, xml_filepath):
        content_tree = etree.parse(xml_filepath)

        for f in content_tree.xpath(".//*[local-name() = 'font']/*[local-name() = 'name']"):
            if f.get('val', '') != '' and self._converter.target_encoding == Encoding.UNICODE:
                f.set('val', helper.replace_font(f.get('val'))[0])

        content_tree.write(xml_filepath)

    def process_file(self):
        self._convert_encoding(os.path.join(self._tmp_dir, 'xl', 'sharedStrings.xml'))
        self._convert_style_file(os.path.join(self._tmp_dir, 'xl', 'styles.xml'))
        self._export_file()


class XlsFileHandler(XlsxFileHandler):
    def __init__(self, filepath, converter, config=None):
        if (config is None) or (config.get(constants.SOFFICE_PATH_KEY, '') == ''):
            raise TypeError('OpenOffice not found, this file will be skipped!')

        self._xlsx_tmp_dir = tempfile.mkdtemp()
        self._config = config

        self._orig_filepath = filepath
        # convert xls to xlsx
        self._temp_filepath = self._convert_file(self._orig_filepath)

        super().__init__(self._temp_filepath, converter)

    def _convert_file(self, filepath):
        filename, extension = os.path.splitext(os.path.basename(filepath))
        convert_to = 'xls' if extension == '.xlsx' else 'xlsx'

        try:
            soffice_path = self._config.get(constants.SOFFICE_PATH_KEY, '')
            subprocess.call(
                [soffice_path, '--headless', '--convert-to', convert_to, filepath, '--outdir', self._xlsx_tmp_dir],
                stdout=open(os.devnull, 'wb')
            )
        except Exception:
            raise errors.FILE_CONVERSION_FAILED

        new_extension = '.' + convert_to
        return os.path.join(self._xlsx_tmp_dir, filename + new_extension)

    def process_file(self):
        super().process_file()

        # convert xlsx to xls
        self._temp_filepath = self._convert_file(self._temp_filepath)
        shutil.copyfile(self._temp_filepath, self._orig_filepath)

    def __del__(self):
        super().__del__()
        try:
            shutil.rmtree(self._xlsx_tmp_dir)
        except Exception:
            pass
