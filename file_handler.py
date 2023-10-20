from lxml import etree
import os
import shutil
import subprocess
import tempfile
import zipfile

from charset import Encoding
import constants
import helper


class FileHandler:
    def __init__(self, filepath, converter, config=None):
        self._filepath = filepath
        self._converter = converter

        self._tmp_dir = tempfile.mkdtemp()
        with open(filepath, 'rb') as f:
            zip_file = zipfile.ZipFile(f)
            zip_file.extractall(self._tmp_dir)
            self._filenames = zip_file.namelist()

    def _convert_encoding(self, xml_filepath):
        content_tree = etree.parse(xml_filepath)
        for r in content_tree.xpath('.//w:r', namespaces=constants.NSMAP):
            t = r.xpath('./w:t', namespaces=constants.NSMAP)
            if t:
                t[0].text, src_encoding = self._converter.convert(t[0].text)
            else:
                continue

            # processing fonts
            if self._converter.target_encoding != Encoding.UNICODE and src_encoding != Encoding.TCVN3:
                continue
            f = r.xpath('./w:rPr/w:rFonts', namespaces=constants.NSMAP)
            if f:
                for attr in f[0].attrib:
                    if src_encoding == Encoding.TCVN3 and f[0].get(attr).endswith('H'):
                        t[0].text = t[0].text.upper()
                    if self._converter.target_encoding == Encoding.UNICODE:
                        f[0].set(attr, helper.replace_font(f[0].get(attr))[0])

        content_tree.write(xml_filepath)

    def _export_file(self):
        with zipfile.ZipFile(self._filepath, 'w') as docx:
            for filename in self._filenames:
                docx.write(os.path.join(self._tmp_dir, filename), filename)

    def process_file(self):
        pass

    def __del__(self):
        if os.path.exists(self._tmp_dir):
            shutil.rmtree(self._tmp_dir)


class DocxFileHandler(FileHandler):
    def process_file(self):
        self._convert_encoding(os.path.join(self._tmp_dir, 'word', 'document.xml'))
        self._export_file()


class DocFileHandler(DocxFileHandler):
    def __init__(self, filepath, converter, config=None):
        if (config is None) or ((config.get(constants.SOFFICE_PATH_KEY, '') == '') and (
           config.get(constants.WORDCONV_PATH_KEY, '') == '')):
            raise TypeError('OpenOffice/M$ Office not found, this file will be skipped!')

        self._docx_tmp_dir = tempfile.mkdtemp()
        self._config = config

        self._orig_filepath = filepath
        print(self._orig_filepath)
        # convert doc to docx
        self._temp_filepath = self._convert_file(self._orig_filepath)
        print(self._temp_filepath)

        super().__init__(self._temp_filepath, converter)

    def _convert_file(self, filepath):
        filename = os.path.basename(filepath)

        if filename.endswith('.doc'):
            new_filepath = os.path.join(
                self._docx_tmp_dir,
                '.docx'.join(os.path.basename(filepath).rsplit('.doc', 1))
            )
            convert_to = 'docx'
        else:
            new_filepath = os.path.join(
                self._docx_tmp_dir,
                '.doc'.join(os.path.basename(filepath).rsplit('.docx', 1))
            )
            convert_to = 'doc'

        soffice_path = self._config.get(constants.SOFFICE_PATH_KEY, '')
        if soffice_path != '':
            subprocess.call(
                [soffice_path, '--headless', '--convert-to', convert_to, filepath, '--outdir', self._docx_tmp_dir],
                stdout=open(os.devnull, 'wb')
            )
            return new_filepath

        # convert docx to doc using wordconv? it won't work
        if filename.endswith('.docx'):
            return filepath

        wordconv_path = self._config.get(constants.WORDCONV_PATH_KEY)
        subprocess.call(
            [wordconv_path, '-oice', '-nme', filepath, new_filepath],
            stdout=open(os.devnull, 'wb')
        )

        return new_filepath

    def process_file(self):
        super().process_file()

        # convert docx to doc
        self._temp_filepath = self._convert_file(self._temp_filepath)
        if os.path.basename(self._temp_filepath).endswith('.doc'):
            shutil.copyfile(self._temp_filepath, self._orig_filepath)
        else:
            os.remove(self._orig_filepath)
            shutil.copyfile(self._temp_filepath, '.docx'.join(self._orig_filepath.rsplit('.doc', 1)))

    def __del__(self):
        super().__del__()
        if os.path.exists(self._docx_tmp_dir):
            shutil.rmtree(self._docx_tmp_dir)


class XlsxFileHandler(FileHandler):
    def _convert_encoding(self, xml_filepath):
        content_tree = etree.parse(xml_filepath)
        filename = os.path.basename(xml_filepath)

        if filename == 'sharedStrings.xml':
            for r in content_tree.xpath(".//*[local-name() = 't']"):
                r.text, _ = self._converter.convert(r.text)
        elif filename == 'styles.xml':
            for f in content_tree.xpath(".//*[local-name() = 'font']/*[local-name() = 'name']"):
                if f.get('val', '') != '' and self._converter.target_encoding == Encoding.UNICODE:
                    f.set('val', helper.replace_font(f.get('val'))[0])
        else:
            pass

        content_tree.write(xml_filepath)

    def process_file(self):
        self._convert_encoding(os.path.join(self._tmp_dir, 'xl', 'sharedStrings.xml'))
        self._convert_encoding(os.path.join(self._tmp_dir, 'xl', 'styles.xml'))
        self._export_file()


class XlsFileHandler(XlsxFileHandler):
    def __init__(self, filepath, converter, config=None):
        if (config is None) or ((config.get(constants.SOFFICE_PATH_KEY, '') == '') and (
           config.get(constants.EXCELCNV_PATH_KEY, '') == '')):
            raise TypeError('OpenOffice/M$ Office not found, this file will be skipped!')

        self._xlsx_tmp_dir = tempfile.mkdtemp()
        self._config = config

        self._orig_filepath = filepath
        # convert xls to xlsx
        self._temp_filepath = self._convert_file(self._orig_filepath)

        super().__init__(self._temp_filepath, converter)

    def _convert_file(self, filepath):
        filename = os.path.basename(filepath)

        if filename.endswith('.xls'):
            new_filepath = os.path.join(
                self._xlsx_tmp_dir,
                '.xlsx'.join(os.path.basename(filepath).rsplit('.xls', 1))
            )
            convert_to = 'xlsx'
        else:
            new_filepath = os.path.join(
                self._xlsx_tmp_dir,
                '.xls'.join(os.path.basename(filepath).rsplit('.xlsx', 1))
            )
            convert_to = 'xls'

        soffice_path = self._config.get(constants.SOFFICE_PATH_KEY, '')
        if soffice_path != '':
            subprocess.call(
                [soffice_path, '--headless', '--convert-to', convert_to, filepath, '--outdir', self._xlsx_tmp_dir],
                stdout=open(os.devnull, 'wb')
            )
            return new_filepath

        # convert xlsx to xls using excelcnv? it won't work
        if filename.endswith('.xlsx'):
            return filepath

        excelcnv_path = self._config.get(constants.EXCELCNV_PATH_KEY)
        subprocess.call(
            [excelcnv_path, '-oice', filepath, new_filepath],
            stdout=open(os.devnull, 'wb')
        )

        return new_filepath

    def process_file(self):
        super().process_file()

        # convert xlsx to xls
        self._temp_filepath = self._convert_file(self._temp_filepath)
        if os.path.basename(self._temp_filepath).endswith('.xls'):
            shutil.copyfile(self._temp_filepath, self._orig_filepath)
        else:
            os.remove(self._orig_filepath)
            shutil.copyfile(self._temp_filepath, '.xlsx'.join(self._orig_filepath.rsplit('.xls', 1)))

    def __del__(self):
        super().__del__()
        if os.path.exists(self._xlsx_tmp_dir):
            shutil.rmtree(self._xlsx_tmp_dir)


class PptxFileHandler(FileHandler):
    def _convert_encoding(self, xml_filepath):
        content_tree = etree.parse(xml_filepath)
        for r in content_tree.xpath('.//a:r', namespaces=constants.NSMAP):
            t = r.xpath('./a:t', namespaces=constants.NSMAP)
            if t:
                t[0].text, src_encoding = self._converter.convert(t[0].text)
            else:
                continue

            # processing fonts
            if self._converter.target_encoding != Encoding.UNICODE and src_encoding != Encoding.TCVN3:
                continue
            f = r.xpath('./a:rPr/a:latin', namespaces=constants.NSMAP)
            if f:
                for attr in f[0].attrib:
                    if src_encoding == Encoding.TCVN3 and f[0].get(attr).endswith('H'):
                        t[0].text = t[0].text.upper()
                    if self._converter.target_encoding == Encoding.UNICODE:
                        f[0].set(attr, helper.replace_font(f[0].get(attr))[0])

            f = r.xpath('./a:rPr/a:cs|./a:rPr/a:latin', namespaces=constants.NSMAP)
            if f:
                orig_typeface = f[0].get('typeface', '')
                if orig_typeface != '':
                    if src_encoding == Encoding.TCVN3 and orig_typeface.endswith('H'):
                        t[0].text = t[0].text.upper()
                    if self._converter.target_encoding == Encoding.UNICODE:
                        typeface, panose = helper.replace_font(f[0].get('typeface'))
                        f[0].set('typeface', typeface)
                        f[0].set('panose', panose)

        content_tree.write(xml_filepath)

    def process_file(self):
        for (root, _, files) in os.walk(os.path.join(self._tmp_dir, 'ppt', 'slides')):
            for file in files:
                self._convert_encoding(os.path.join(root, file))
        self._export_file()


class PptFileHandler(PptxFileHandler):
    def __init__(self, filepath, converter, config=None):
        if (config is None) or ((config.get(constants.SOFFICE_PATH_KEY, '') == '') and (
           config.get(constants.PPCNVCOM_PATH_KEY, '') == '')):
            raise TypeError('OpenOffice/M$ Office not found, this file will be skipped!')

        self._pptx_tmp_dir = tempfile.mkdtemp()
        self._config = config

        self._orig_filepath = filepath
        # convert xls to xlsx
        self._temp_filepath = self._convert_file(self._orig_filepath)

        super().__init__(self._temp_filepath, converter)

    def _convert_file(self, filepath):
        filename = os.path.basename(filepath)

        if filename.endswith('.ppt'):
            new_filepath = os.path.join(
                self._pptx_tmp_dir,
                '.pptx'.join(os.path.basename(filepath).rsplit('.ppt', 1))
            )
            convert_to = 'pptx'
        else:
            new_filepath = os.path.join(
                self._pptx_tmp_dir,
                '.ppt'.join(os.path.basename(filepath).rsplit('.pptx', 1))
            )
            convert_to = 'ppt'

        soffice_path = self._config.get(constants.SOFFICE_PATH_KEY, '')
        if soffice_path != '':
            subprocess.call(
                [soffice_path, '--headless', '--convert-to', convert_to, filepath, '--outdir', self._pptx_tmp_dir],
                stdout=open(os.devnull, 'wb')
            )
            return new_filepath

        # convert pptx to ppt using ppconvcom? it won't work
        if filename.endswith('.pptx'):
            return filepath

        excelcnv_path = self._config.get(constants.EXCELCNV_PATH_KEY)
        subprocess.call(
            [excelcnv_path, '-oice', filepath, new_filepath],
            stdout=open(os.devnull, 'wb')
        )

        return new_filepath

    def process_file(self):
        super().process_file()

        # convert pptx to ppt
        self._temp_filepath = self._convert_file(self._temp_filepath)
        if os.path.basename(self._temp_filepath).endswith('.ppt'):
            shutil.copyfile(self._temp_filepath, self._orig_filepath)
        else:
            os.remove(self._orig_filepath)
            shutil.copyfile(self._temp_filepath, '.pptx'.join(self._orig_filepath.rsplit('.ppt', 1)))

    def __del__(self):
        super().__del__()
        if os.path.exists(self._pptx_tmp_dir):
            shutil.rmtree(self._pptx_tmp_dir)
