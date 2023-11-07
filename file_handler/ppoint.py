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


class PptxFileHandler(FileHandler):
    def _convert_encoding(self, xml_filepath):
        content_tree = etree.parse(xml_filepath)
        for r in content_tree.xpath('.//a:r', namespaces=constants.NSMAP):
            text_fields = r.xpath('./a:t', namespaces=constants.NSMAP)
            for t in text_fields:
                t.text, src_encoding = self._converter.convert(t.text)

            # processing fonts
            if self._converter.target_encoding != Encoding.UNICODE and src_encoding != Encoding.TCVN3:
                continue

            font_fields = r.xpath('./a:rPr/a:cs|./a:rPr/a:latin', namespaces=constants.NSMAP)
            for f in font_fields:
                orig_typeface = f.get('typeface')
                if self._converter.target_encoding == Encoding.UNICODE:
                    typeface, panose = helper.replace_font(orig_typeface)
                    f.set('typeface', typeface)
                    f.set('panose', panose)
                if src_encoding == Encoding.TCVN3 and orig_typeface.endswith('H'):
                    for t in text_fields:
                        t.text = t.text.upper()

        content_tree.write(xml_filepath)

    def process_file(self):
        for (root, _, files) in os.walk(os.path.join(self._tmp_dir, 'ppt', 'slides')):
            for file in files:
                self._convert_encoding(os.path.join(root, file))
            break
        self._export_file()


class PptFileHandler(PptxFileHandler):
    def __init__(self, filepath, converter, config=None):
        if (config is None) or (config.get(constants.SOFFICE_PATH_KEY, '') == ''):
            raise TypeError('OpenOffice/M$ Office not found, this file will be skipped!')

        self._pptx_tmp_dir = tempfile.mkdtemp()
        self._config = config

        self._orig_filepath = filepath
        # convert ppt to pptx
        self._temp_filepath = self._convert_file(self._orig_filepath)

        super().__init__(self._temp_filepath, converter)

    def _convert_file(self, filepath):
        filename, extension = os.path.splitext(os.path.basename(filepath))
        convert_to = 'ppt' if extension == '.pptx' else 'pptx'

        try:
            soffice_path = self._config.get(constants.SOFFICE_PATH_KEY, '')
            subprocess.call(
                [soffice_path, '--headless', '--convert-to', convert_to, filepath, '--outdir', self._pptx_tmp_dir],
                stdout=open(os.devnull, 'wb')
            )
        except Exception:
            raise errors.FILE_CONVERSION_FAILED

        new_extension = '.' + convert_to
        return os.path.join(self._pptx_tmp_dir, filename + new_extension)

    def process_file(self):
        super().process_file()

        # convert pptx to ppt
        self._temp_filepath = self._convert_file(self._temp_filepath)
        shutil.copyfile(self._temp_filepath, self._orig_filepath)

    def __del__(self):
        super().__del__()
        try:
            shutil.rmtree(self._pptx_tmp_dir)
        except Exception:
            pass
