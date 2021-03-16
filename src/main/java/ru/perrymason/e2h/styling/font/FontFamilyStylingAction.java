package ru.perrymason.e2h.styling.font;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import ru.perrymason.e2h.Excel2Html;

public class FontFamilyStylingAction extends FontStylingAction {

    public FontFamilyStylingAction(FontResolver fontResolver) {
        super(fontResolver);
    }

    @Override
    public void perform(Cell cell, CellStyle cellStyle, Excel2Html.CellSpans cellSpans, StringBuilder htmlStyle) {
        Font font = fontResolver.getFont(cellStyle);
        htmlStyle.append("font-family:").append(getFontFamilies(font)).append(";");
    }

}
