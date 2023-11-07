import argparse
import colorama

from core.handler import Handler


def main():
    parser = argparse.ArgumentParser(
        description='Convert between Vietnamese encoding formats.\n' +
                    'Supported encoding:\n' +
                    '\tUNICODE\n' +
                    '\tTCVN3\n' +
                    '\tVNI\n' +
                    '\tVIQR\n' +
                    '\tVISCII',
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('input', help='The string or path to file/folder that need to be processed')
    parser.add_argument('target_encoding', help='Convert text to this encoding')

    parser.add_argument('--source-encoding', action='store', default='', help='Specify the original encoding. ' +
                        'If not specified, the program will try to determine original encoding automatically')
    parser.add_argument('--input-type', action='store', default='', help='Specify the type for the passed input. ' +
                        'If not specified, the program will try to detect the type automatically. Supported type: ' +
                        'string, file, directory')
    args = parser.parse_args()

    colorama.init()
    handler = Handler(
        input=args.input,
        target_encoding=args.target_encoding,
        input_type=args.input_type if args.input_type != '' else None,
        source_encoding=args.source_encoding if args.source_encoding != '' else None,
    )
    handler.process()


if __name__ == "__main__":
    main()
