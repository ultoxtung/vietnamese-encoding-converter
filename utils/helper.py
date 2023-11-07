import constants.constants as constants


def create_new_file_path(path):
    f = path.rsplit('.', 1)
    return f[0] + '_old.' + f[1]


def create_new_dir_path(path):
    f = path.rstrip('/\\')
    return f + '_old'


def replace_font(current_font: str):
    font_lower = current_font.lower()
    for f in constants.FONTS:
        if f.typeface.lower() in font_lower:
            return f.typeface, f.panose

    return constants.FONTS[0].typeface, constants.FONTS[1].panose
