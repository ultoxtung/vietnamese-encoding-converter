import yaml
import colorama
import os
import platform
import shutil

import constants
import converter
import file_handler
import helper


class Handler:
    def __init__(self, input, target_encoding, input_type=None, source_encoding=None):
        self.input_type = input_type
        if self.input_type is None:
            # detect input type
            self.input_type = constants.InputType.STRING
            if os.path.exists(input):
                self.input_type = constants.InputType.DIRECTORY if os.path.isdir(input)\
                    else constants.InputType.FILE

        self._has_old_office_files = False
        if self.input_type == constants.InputType.STRING:
            self.source_string = input
        elif self.input_type == constants.InputType.FILE:
            if not self.__check_file_is_supported(input):
                raise TypeError('File is not supported!')
            self.paths = [input]
            shutil.copy2(input, helper.create_new_file_path(input))
        else:
            self.paths = []
            for (root, _, files) in os.walk(input):
                self.paths.extend([os.path.join(root, file) for file in files if self.__check_file_is_supported(file)])
            shutil.copytree(input, helper.create_new_dir_path(input), dirs_exist_ok=True)

        self.converter = converter.Converter(source_encoding, target_encoding)

    def __check_file_is_supported(self, path):
        filename = os.path.basename(path)
        if filename.endswith(('.doc', '.xls', '.ppt')):
            self._has_old_office_files = True
            return True
        return os.path.basename(path).endswith(('.docx', '.xlsx', '.pptx'))

    def __switch_file_handler(self, filepath):
        filename = os.path.basename(filepath)
        if filename.endswith('.doc'):
            return file_handler.DocFileHandler(filepath, self.converter, self._config)
        elif filename.endswith('.docx'):
            return file_handler.DocxFileHandler(filepath, self.converter)
        elif filename.endswith('.xls'):
            return file_handler.XlsFileHandler(filepath, self.converter, self._config)
        elif filename.endswith('.xlsx'):
            return file_handler.XlsxFileHandler(filepath, self.converter)
        elif filename.endswith('.ppt'):
            return file_handler.PptFileHandler(filepath, self.converter, self._config)
        elif filename.endswith('.pptx'):
            return file_handler.PptxFileHandler(filepath, self.converter)
        else:
            return None

    def __load_config_file(self):
        if os.path.exists(constants.CONFIG_FILE):
            with open(constants.CONFIG_FILE, 'r') as config_file:
                self._config = yaml.load(config_file, Loader=yaml.FullLoader)
                soffice_path = self._config.get(constants.SOFFICE_PATH_KEY, '')
                if soffice_path != '' and shutil.which(soffice_path) is not None:
                    return

        print('[INFO] Trying to detect OpenOffice installation...')
        self._config = {}
        if platform.system() == 'Windows':
            for (root, _, files) in os.walk('C:\\'):
                for file in files:
                    if file.lower() == 'soffice.exe':
                        self._config[constants.SOFFICE_PATH_KEY] = os.path.join(root, file)
        else:
            if shutil.which('soffice') is not None:
                self._config[constants.SOFFICE_PATH_KEY] = 'soffice'

        with open(constants.CONFIG_FILE, 'w') as config_file:
            yaml.dump(self._config, config_file)

    def process(self):
        if self.input_type == constants.InputType.STRING:
            print(self.converter.convert(self.source_string)[0])
            return

        if self._has_old_office_files:
            self.__load_config_file()

        for filepath in self.paths:
            print('Processing file "{}" ...'.format(filepath))
            try:
                file_handler = self.__switch_file_handler(filepath)
                if file_handler is None:
                    raise TypeError('File is not supported')

                file_handler.process_file()
            except Exception as e:
                print('\t', colorama.Fore.RED, '[ERROR]', e, colorama.Style.RESET_ALL)
            else:
                print('\t', colorama.Fore.GREEN, 'OK', colorama.Style.RESET_ALL)
