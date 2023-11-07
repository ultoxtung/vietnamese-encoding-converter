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


class DocxFileHandler(FileHandler):
    def _convert_encoding(self, xml_filepath):
        content_tree = etree.parse(xml_filepath)

        for r in content_tree.xpath('.//w:r', namespaces=constants.NSMAP):
            text_fields = r.xpath('./w:t', namespaces=constants.NSMAP)
            for t in text_fields:
                t.text, src_encoding = self._converter.convert(t.text)

            # processing fonts
            if self._converter.target_encoding != Encoding.UNICODE and src_encoding != Encoding.TCVN3:
                continue
            font_fields = r.xpath('./w:rPr/w:rFonts', namespaces=constants.NSMAP)
            for f in font_fields:
                for attr in f.attrib:
                    font = f.get(attr)
                    if self._converter.target_encoding == Encoding.UNICODE:
                        f.set(attr, helper.replace_font(font)[0])
                    if src_encoding == Encoding.TCVN3 and font.endswith('H'):
                        for t in text_fields:
                            t.text = t.text.upper()

        content_tree.write(xml_filepath)

    def _convert_style_file(self, xml_filepath):
        if self._converter.target_encoding != Encoding.UNICODE:
            return

        content_tree = etree.parse(xml_filepath)

        for f in content_tree.xpath('.//w:rPr/w:rFonts', namespaces=constants.NSMAP):
            for attr in f.attrib:
                if self._converter.target_encoding == Encoding.UNICODE:
                    f.set(attr, helper.replace_font(f.get(attr))[0])

        content_tree.write(xml_filepath)

    def process_file(self):
        self._convert_encoding(os.path.join(self._tmp_dir, 'word', 'document.xml'))
        self._convert_style_file(os.path.join(self._tmp_dir, 'word', 'styles.xml'))
        self._export_file()


class DocFileHandler(DocxFileHandler):
    def __init__(self, filepath, converter, config=None):
        if (config is None) or (config.get(constants.SOFFICE_PATH_KEY, '') == ''):
            raise errors.OPENOFFICE_NOT_FOUND

        self._docx_tmp_dir = tempfile.mkdtemp()
        self._config = config

        self._orig_filepath = filepath
        # convert doc to docx
        self._temp_filepath = self._convert_file(self._orig_filepath)

        super().__init__(self._temp_filepath, converter)

    def _convert_file(self, filepath):
        filename, extension = os.path.splitext(os.path.basename(filepath))
        convert_to = 'doc' if extension == '.docx' else 'docx'

        try:
            soffice_path = self._config.get(constants.SOFFICE_PATH_KEY)
            subprocess.call(
                [soffice_path, '--headless', '--convert-to', convert_to, filepath, '--outdir', self._docx_tmp_dir],
                stdout=open(os.devnull, 'wb')
            )
        except Exception:
            raise errors.FILE_CONVERSION_FAILED

        new_extension = '.' + convert_to
        return os.path.join(self._docx_tmp_dir, filename + new_extension)

    def process_file(self):
        super().process_file()

        # convert docx to doc
        self._temp_filepath = self._convert_file(self._temp_filepath)
        shutil.copyfile(self._temp_filepath, self._orig_filepath)

    def __del__(self):
        super().__del__()
        try:
            shutil.rmtree(self._docx_tmp_dir)
        except Exception:
            pass
