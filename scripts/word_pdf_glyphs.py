"""Physical PDF glyph extraction shared by native Word comparison gates."""
import ctypes
import math
import pypdfium2_raw as raw


def glyphs(page, no_authored_spaces):
    result = []
    text = page.get_textpage()
    try:
        for index in range(text.count_chars()):
            char = chr(raw.FPDFText_GetUnicode(text, index))
            generated = raw.FPDFText_IsGenerated(text, index) == 1
            if generated and (char in '\r\n' or (char == ' ' and no_authored_spaces)):
                continue
            if generated and char == ' ':
                # PDFium reconstructs Chromium's word gaps as spaces. They have
                # no painted glyph, font matrix or fill color. Keep the text
                # identity; measure the gap through its neighboring glyphs.
                result.append({'text': char, 'generated': True})
                continue
            x, y = ctypes.c_double(), ctypes.c_double()
            matrix = raw.FS_MATRIX()
            rgba = [ctypes.c_uint() for _ in range(4)]
            assert raw.FPDFText_GetCharOrigin(text, index, x, y)
            assert raw.FPDFText_GetMatrix(text, index, ctypes.byref(matrix))
            assert raw.FPDFText_GetFillColor(text, index, *rgba)
            result.append({
                'text': char, 'x': x.value, 'y': page.get_height() - y.value,
                'size': raw.FPDFText_GetFontSize(text, index) * math.hypot(matrix.c, matrix.d),
                'color': [channel.value for channel in rgba],
            })
    finally:
        text.close()
    return result
