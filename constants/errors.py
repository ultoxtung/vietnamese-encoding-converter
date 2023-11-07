# File type and conversion
FILE_NOT_SUPPORTED = TypeError('File is not supported!')
OPENOFFICE_NOT_FOUND = TypeError('OpenOffice not found, this file will be skipped!')
FILE_CONVERSION_FAILED = TypeError('File format conversion failed, this file will be skipped!')

# Encoding validation
SOURCE_ENCODING_NOT_SUPPORTED = TypeError('Source encoding is not supported!')
TARGET_ENCODING_NOT_SUPPORTED = TypeError('Target encoding is not supported!')
SAME_TARGET_ENCODING_AS_SOURCE = TypeError('Target encoding is the same as source encoding!')

# Info messages
MSG_DETECTING_OPENOFFICE = '[INFO] Trying to detect OpenOffice installation...'
