from openpyxl.styles import Border, Side, Alignment, Font

BOLD_FONT = Font("Calibri", 16, bold=True)
FONT = Font("Calibri", 16)
CENTER_ALIGNMENT = Alignment("center", "center")
BORDER = Border(left=Side(border_style="thin",
                          color='FF000000'),
                right=Side(border_style="thin",
                           color='FF000000'),
                top=Side(border_style="thin",
                         color='FF000000'),
                bottom=Side(border_style="thin",
                            color='FF000000')
                )