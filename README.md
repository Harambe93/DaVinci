# DaVinSheet

Converts any image into an excel spreadsheet, so that every cell of the spreadsheet
is now a pixel.

![demonstration](/img/demo.png)

### Usage

Run

```sh
davinsheet.py /path/to/my/image.jpg
```

on any image file to generate a very ... artistic excel spreadsheet.

### Info

The output image gets downsized so your spreadsheet software doesn't crash when
asked to load a couple of thousand tiny coloured cells.

Tested using LibreOffice Calc, cell sizes might be off in other spreadsheet softwares.
The cell settings can be easily tweaked in the script.

### Dependencies

Python 3.4 or above, [openpyxl](https://pypi.org/project/openpyxl/)
