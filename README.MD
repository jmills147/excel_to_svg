# Excel to SVG
A small no-installation Windows application to export a chart from Excel to an svg file. Saves messing around with pdfs and Inkscape.

Alternatively a Python user can run the underlying script directly.

## Windows .exe

* [Download](https://1drv.ms/u/s!AlQAyjpM4rC5kbt8IbYKyDkj7f9XCA?e=rlxdhY) from OneDrive (20mb)

* No installation. Just download and launch.

![screenshot](screenshot_2.png)

## Python script

1. Install requirements:
```
pip install -r requirements.txt
```
2. In Excel, select the chart you wish to export.

3. Run:
```
python svg_from_pdf.py
```

## Range export

Although this application/script is intended for charts, it works for any range (which may contain shapes, text boxes, etc.).

Just set the print area and make sure that everything you are interested in is on the first page (Page Setup – Scaling – Fit to 1 page wide / tall).

The svg will include headers and footers so remove them in Excel if not wanted.

## Issues
#### Problem installing xlWings
If you get the following error when running the script:

```
ImportError: DLL load failed while importing win32api: The specified procedure could not be found.
```
try reverting the pywin32 version:
```
pip install pywin32==300
```
(See [stackoverflow](https://stackoverflow.com/questions/40981120/python-pip-cannot-find-pywin32-on-windows) disussion)
