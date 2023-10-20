def create_new_file_path(path):
    f = path.rsplit('.', 1)
    return f[0] + '_old.' + f[1]


def create_new_dir_path(path):
    f = path.rstrip('/\\')
    return f + '_old'


def replace_font(current_font: str):
    if 'arial' in current_font.lower():
        return 'Arial', '02110604020202020204'
    if 'tahoma' in current_font.lower():
        return 'Tahoma', '02110604030504040204'
    if 'calibri' in current_font.lower():
        return 'Calibri', '02150502020204030204'
    return 'Times New Roman', '02020603050405020304'
