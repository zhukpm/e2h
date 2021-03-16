package ru.perrymason.e2h.styling.font;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import ru.perrymason.e2h.Excel2Html;

public class FontStyleStylingAction extends FontStylingAction {

    public FontStyleStylingAction(FontResolver fontResolver) {
        super(fontResolver);
    }

    @Override
    public void perform(Cell cell, CellStyle cellStyle, Excel2Html.CellSpans cellSpans, StringBuilder htmlStyle) {
        Font font = fontResolver.getFont(cellStyle);
        String str = getItalic(font);
        if (str.length() > 0) {
            htmlStyle.append("font-style:").append(str).append(";");
        }
        str = getBold(font);
        if (str.length() > 0) {
            htmlStyle.append("font-weight:").append(str).append(";");
        }
        if (font.getUnderline() != Font.U_NONE) {
            htmlStyle.append("text-decoration:underline;");
        }
    }

}
