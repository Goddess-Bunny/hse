def create_format(wb, font_color=None, bg_color=None, font_name='Times New Roman',
                  font_size=12, valign='vcenter', halign='center', border=1, bold=False,
                  text_wrap=True, italic=False, rotation=0, indent=0):
    _format = wb.add_format({'border': border, 'bold': bold, 'text_wrap': text_wrap})
    _format.set_font_name(font_name)
    _format.set_font_size(font_size)
    _format.set_italic(italic)
    _format.set_align(valign)
    _format.set_align(halign)
    _format.set_rotation(rotation)
    _format.set_indent(indent)
    if font_color is not None:
        _format.set_font_color(font_color)
    if bg_color is not None:
        _format.set_bg_color(bg_color)

    return _format


def process_day(day):
    if 10 <= day <= 20:
        return 'дней'
    elif day % 10 == 1:
        return 'день'
    elif 1 < day % 10 < 5:
        return 'дня'
    else:
        return 'дней'


class RowWriter:
    def __init__(self, workbook, worksheet, row=3, col=0, height=1):
        self.height = height
        self.cur_col = col
        self.first_row = row
        self.ws = worksheet
        self.wb = workbook

    def add_cols(self, content, style_mask, width_mask):
        if type(content[0]) is not list:
            content, style_mask, width_mask = [content], [style_mask], [width_mask]

        height_to_merge = 0
        for row in range(len(content)):
            if row == len(content) - 1:
                height_to_merge = self.height - len(content)

            col = 0
            for name, style, width in zip(content[row], style_mask[row], width_mask[row]):
                if height_to_merge == 0 and width == 1:
                    self.ws.write(self.first_row + row, self.cur_col + col, name, style)
                else:
                    self.ws.merge_range(self.first_row + row, self.cur_col + col,
                                        self.first_row + row + height_to_merge, self.cur_col + col + width - 1,
                                        name, style)
                col += width

        self.cur_col += sum(width_mask[0])

    def insert_green(self):
        green = self.wb.add_format()
        green.set_bg_color("#009933")

        if self.height > 1:
            self.ws.merge_range(self.first_row, self.cur_col,
                                self.first_row + self.height - 1, self.cur_col,
                                "", green
                                )
        else:
            self.ws.write(self.first_row, self.cur_col, "", green)

        self.ws.set_column(self.cur_col, self.cur_col, 2)

        self.cur_col += 1

