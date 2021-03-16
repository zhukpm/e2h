package ru.perrymason.e2h.styling.font;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import ru.perrymason.e2h.Excel2Html;

public class FullBundledFontStylingAction extends FontStylingAction {

    public FullBundledFontStylingAction(FontResolver fontResolver) {
        super(fontResolver);
    }

    @Override
    public void perform(Cell cell, CellStyle cellStyle, Excel2Html.CellSpans cellSpans, StringBuilder htmlStyle) {
        Font font = fontResolver.getFont(cellStyle);
        htmlStyle.append("font:").append(getItalic(font)).append(getBold(font))
                .append(font.getFontHeightInPoints()).append("pt ").append(getFontFamilies(font)).append(";");
        if (font.getUnderline() != Font.U_NONE) {
            htmlStyle.append("text-decoration:underline;");
        }
    }

}
