"""Print layout for the offering report.

The report is short, so on paper it sits in the top half of the page with a
wide band of white space underneath. Growing the row heights to fill the
printable area removes that band without changing any content.
"""

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

POINTS_PER_INCH = 72.0
DEFAULT_ROW_HEIGHT = 15.0  # Excel's default, used when a row has no explicit height


def apply_print_layout(worksheet, *, paper_height_in: float = 11.0,
                       top_margin_in: float = 0.25, bottom_margin_in: float = 0.25,
                       left_margin_in: float = 0.25, right_margin_in: float = 0.25,
                       header_in: float = 0.0, footer_in: float = 0.0,
                       first_row: int = 1, last_row: int = None,
                       min_row_height: float = DEFAULT_ROW_HEIGHT,
                       max_row_height: float = 45.0,
                       fill_ratio: float = 1.0,
                       set_print_area: bool = True) -> dict:
    """Shrink page margins and grow row heights so the report fills the page.

    Returns a small dict describing what it did, which makes the effect
    testable without opening Excel.
    """
    last_row = last_row or worksheet.max_row
    if last_row < first_row:
        return {"rows_scaled": 0, "scale": 1.0, "available_points": 0.0}

    worksheet.page_margins.top = top_margin_in
    worksheet.page_margins.bottom = bottom_margin_in
    worksheet.page_margins.left = left_margin_in
    worksheet.page_margins.right = right_margin_in
    worksheet.page_margins.header = header_in
    worksheet.page_margins.footer = footer_in

    worksheet.page_setup.orientation = "portrait"
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 1
    # fitToWidth/fitToHeight are ignored unless this property is set too.
    worksheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    worksheet.print_options.horizontalCentered = True
    worksheet.print_options.verticalCentered = False

    if set_print_area and not worksheet.print_area:
        # Only set a print area when the sheet does not already define one.
        # The production template prints $B$1:$L$32 deliberately - column A
        # holds a duplicate of the 수표정리 amounts in column B, and widening
        # the range to A would put that second column of numbers on the paper.
        # Build the reference from the column index directly: cell() returns a
        # MergedCell for merged ranges, which has no column_letter.
        last_col = get_column_letter(worksheet.max_column)
        worksheet.print_area = f"A{first_row}:{last_col}{last_row}"

    usable_in = paper_height_in - top_margin_in - bottom_margin_in - header_in - footer_in
    available_points = usable_in * POINTS_PER_INCH * fill_ratio

    heights = {}
    for row_number in range(first_row, last_row + 1):
        existing = worksheet.row_dimensions[row_number].height
        heights[row_number] = existing if existing else DEFAULT_ROW_HEIGHT

    current_total = sum(heights.values())
    if current_total <= 0:
        return {"rows_scaled": 0, "scale": 1.0, "available_points": available_points}

    scale = available_points / current_total
    if scale <= 1.0:
        # Already fills the page (or overflows) - fitToHeight handles the rest.
        return {"rows_scaled": 0, "scale": scale, "available_points": available_points}

    for row_number, height in heights.items():
        scaled = max(min_row_height, min(max_row_height, height * scale))
        worksheet.row_dimensions[row_number].height = round(scaled, 2)

    return {
        "rows_scaled": len(heights),
        "scale": round(scale, 3),
        "available_points": round(available_points, 1),
        "new_total_points": round(sum(worksheet.row_dimensions[r].height for r in heights), 1),
    }
