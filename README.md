# Chuyển đổi qua lại giữa một số bảng mã tiếng Việt

## Các bảng mã được hỗ trợ:

- UNICODE
- TCVN3
- VIQR
- VNI
- VISCII

## Requirement:

- Python 3.x
- OpenOffice (Không cần nếu không convert các file Office 97-2003 (.doc, .xls, .ppt))
- `lxml`, `pyyaml` và `colorama`

## Sử dụng:

```
usage: main.py [-h] [--source-encoding SOURCE_ENCODING] [--input-type INPUT_TYPE] input target_encoding

Convert between Vietnamese encoding formats.
Supported encoding:
        UNICODE
        TCVN3
        VNI
        VIQR
        VISCII

positional arguments:
  input                 The string or path to file/folder that need to be processed
  target_encoding       Convert text to this encoding

optional arguments:
  -h, --help            show this help message and exit
  --source-encoding SOURCE_ENCODING
                        Specify the original encoding. If not specified, the program will try to determine original encoding automatically
  --input-type INPUT_TYPE
                        Specify the type for the passed input. If not specified, the program will try to detect the type automatically. Supported type: string, file, directory
```

## Ví dụ:

```
~ $ python main.py "ThËt kh«ng thÓ tin næi, thËt lµ tuyÖt vêi" unicode --source-encoding tcvn3

Thật không thể tin nổi, thật là tuyệt vời
```

```
~ $ python main.py "/home/you/Documents" unicode --input-type directory

Processing file "/home/you/Documents/book1.xlsx" ...
          OK
Processing file "/home/you/Documents/doc1.docx" ...
          OK
Processing file "/home/you/Documents/doc2.doc" ...
          OK
```

## Known issue:

- Một số font unicode có thể bị chuyển về Times New Roman khi chuyển file sang unicode
- Khi chuyển từ VISCII sang các bảng mã khác có một số chỗ bị lẫn lộn hoa thường

## ~~Where is unit test?~~

~~Trust me bro™~~