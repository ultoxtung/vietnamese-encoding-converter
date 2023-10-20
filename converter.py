from charset import Charset, Encoding


class Converter:
    def __init__(self, source_encoding, target_encoding):
        if source_encoding is not None:
            self.source_encoding = self.__get_encoding(source_encoding)
            if self.source_encoding is None:
                raise TypeError('Source encoding is not supported!')
        else:
            self.source_encoding = None

        if target_encoding is not None:
            self.target_encoding = self.__get_encoding(target_encoding)
            if self.target_encoding is None:
                raise TypeError('Target encoding is not supported!')
        else:
            raise TypeError('Target encoding is not supported!')

        if self.target_encoding == self.source_encoding:
            raise TypeError('Target encoding is the same as source encoding!')

        self.__populate_converter_table()

    def __get_encoding(self, encoding):
        for e in Encoding:
            if e.name.casefold() == encoding.casefold():
                return e

        return None

    def __populate_converter_table(self):
        def make_charset_map(source_charset, target_charset):
            charset_map = {}
            for i in range(len(source_charset)):
                charset_map[source_charset[i]] = target_charset[i]
            return charset_map

        charset = Charset()
        self.converter_table = [[None for _ in range(len(Encoding))] for _ in range(len(Encoding))]

        target_charset = charset.get_charset(self.target_encoding)
        if self.source_encoding is not None:
            source_charset = charset.get_charset(self.source_encoding)
            self.converter_table[self.source_encoding.value][self.target_encoding.value] = make_charset_map(
                source_charset, target_charset)
            self.converter_table[self.target_encoding.value][self.target_encoding.value] = make_charset_map(
                target_charset, target_charset)
        else:
            for encoding in Encoding:
                source_charset = charset.get_charset(encoding)
                self.converter_table[encoding.value][self.target_encoding.value] = make_charset_map(
                    source_charset, target_charset)

    def __convert_from(self, src_string, source_encoding):
        i = 0
        dest_string = ''
        replacement_count = 0

        while i < len(src_string):
            replaced = False
            for lg in range(1, 4):
                char = self.converter_table[source_encoding.value][self.target_encoding.value].get(src_string[i:i + lg])
                if char is not None:
                    dest_string += char
                    replacement_count += 1
                    i += lg
                    replaced = True
                    break

            if not replaced:
                dest_string += src_string[i:i + 1]
                i += 1

        return dest_string, replacement_count

    def convert(self, src_string):
        if not len(src_string):
            return

        result_string, replacement_count = self.__convert_from(src_string, self.target_encoding)
        ratio = replacement_count / len(src_string)
        src_encoding = self.target_encoding

        if self.source_encoding is not None:
            dest_string, replacement_count = self.__convert_from(src_string, self.source_encoding)
            if len(dest_string) and (replacement_count / len(dest_string)) > ratio:
                result_string = dest_string
                ratio = replacement_count / len(dest_string)
                src_encoding = self.source_encoding
        else:
            for encoding in Encoding:
                if encoding.value == self.target_encoding.value:
                    continue
                dest_string, replacement_count = self.__convert_from(src_string, encoding)
                if len(dest_string) and (replacement_count / len(dest_string)) > ratio:
                    result_string = dest_string
                    ratio = replacement_count / len(dest_string)
                    src_encoding = encoding

        return result_string, src_encoding
