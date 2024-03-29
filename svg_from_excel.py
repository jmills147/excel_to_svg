"""Simple script to produce svg from Excel chart

Example:
    Select a chart in the active sheet in Excel
    Run this script:

        $ python svg_from_excel.py

    Produces svg in the same folder as workbook.

    Also works if a range is selected, where the svg is Page 1 of the Print
    Range, including headers and footers. You may want to 'Fit to 1 Page Wide/Tall'
    and remove headers and footers to get what you want.

    There is a simple GUI that calls this script – gui.py – and a no-install
    windows executable.

"""
import xlwings as xw
from pathlib import Path
import fitz
import re

def generate_pdf_path(sht: xw.Sheet) -> Path:

    """Generates a full path for the target pdf file based on active
    chart in the selected sheet

    Given by:
        {{ excel book folder* }} \ {{ book_name without suffix }}{{ sheet name }}_{{ chart name }}

    Spaces in the book name, sheet name or chart name are replaced with underscores.

    If the book does not have a folder – because it has not been saved –
    then use the user's home folder.

    If there is no active chart in the sheet then return:
        {{ excel book folder* }} \ {{ book_name }}{{ sheet name }}_{{ chart name }}

    Args:
        sht: xlWings sheet object

    Returns:
        Full path save-to location for pdf

    """
    wb = sht.book
    
    try:
        sheet_and_chart = wb.api.ActiveChart.Name
    except AttributeError:
        print('\n*** No chart is selected in Excel. Exporting the range instead. ***\n')
        sheet_and_chart = sht.name
    
    book_name_stub = wb.name.split('.')[0] # i.e. name without '.xlsx' suffix
    book_sheet_chart = book_name_stub + ' ' + sheet_and_chart
    book_sheet_chart = book_sheet_chart.replace(' ', '_')
    
    pdf_filename = book_sheet_chart + '.pdf'

    wbk_path = Path(wb.fullname)
    
    if wbk_path.is_file():
        pdf_path = wbk_path.parent / pdf_filename
    else:  # workbook has not been saved, so save pdf to user home directory
        pdf_path = Path.home() / pdf_filename
    
    return pdf_path


def svg_from_excel_pdf(pdf_path: Path,
                       svg_path: Path) -> int:

    """Extracts svg from pdf and saves it to save path (but with .svg suffix)

    Args:
        pdf_path:
        svg_path:

    Returns:
        0 fof success

    """

    with fitz.open(pdf_path) as pdf_file:
        svg = pdf_file[0].get_svg_image()
    
    svg = reduce_svg_viewbox(svg)
    
    with open(svg_path, 'wt') as svg_file:
        svg_file.write(svg)
    
    return 0


def reduce_svg_viewbox(svg: str) -> str:

    """Removes the padding around svg

    Args:
        svg: The svg string, usually for svg extracted from pdf from Excel

    Returns:
        svg string with the viewBox changed to remove the padding

    """

    x_max = 0
    x_min = 99999
    y_max = 0
    y_min = 99999

    # don't want clip paths with url before id e.g. clip-path="url(#cp0)">
    for mtch in re.finditer('\<clipPath id="cp.*?\n<path transform.*? d="(.*?)"', svg):
            
        s = mtch.group(1)
        # print(s)
        for i, mtch2 in enumerate(re.finditer('[L|M] (\S+) (\S+)', s)):
            try:
                x = float(mtch2.group(1))
                y = float(mtch2.group(2))
            except ValueError:
                print(s, mtch2.group(1), mtch2.group(2))
                raise Exception
            if x < x_min:
                x_min = x
            if x > x_max:
                x_max = x
            if y < y_min:
                y_min = y
            if y > y_max:
                y_max = y

    y_edge = re.search(r'"matrix\(.*,(.*)\)"', svg).group(1)  # multiple matches but all the same: take first one
    y_edge = float(y_edge)
    
    # print(f'{x_min=}, {x_max=}, {y_min=}, {y_max=}, {y_edge=}')

    reduced_viewbox_string= f'{x_min:.2f} {y_edge-y_max:.2f} {x_max-x_min:.2f} {y_max-y_min:.2f}'
    # print(reduced_viewbox_string)
    svg = re.sub(r'viewBox=".*?"', f'viewBox="{reduced_viewbox_string}"', svg, count=1)  # sub first match at top only

    # remove width and height
    svg = re.sub(r' width=".*?" height=".*?"', '', svg, count=1)

    return svg


def export_active_chart_to_svg() -> str:

    """Main routine

    Generate save-to path
    Excel export to pdf
    Extract svg from pdf
    Delete pdf

    Args:

    Returns:
        String path to svg, for gui

    """

    sht = xw.sheets.active
    pdf_path = generate_pdf_path(sht)

    svg_path = pdf_path.with_suffix('.svg')

    sht.api.ExportAsFixedFormat(0, str(pdf_path))
    
    svg_from_excel_pdf(pdf_path, svg_path)

    pdf_path.unlink()  # delete pdf file

    print(f'svg produced: {svg_path}')

    return str(svg_path)

0
if __name__ == '__main__':

    _ = export_active_chart_to_svg()
