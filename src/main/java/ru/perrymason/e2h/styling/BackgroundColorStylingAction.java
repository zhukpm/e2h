package ru.perrymason.e2h.styling;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import ru.perrymason.e2h.Excel2Html;

public abstract class BackgroundColorStylingAction implements StylingAction {
    @Override
    public void perform(Cell cell, CellStyle cellStyle, Excel2Html.CellSpans cellSpans, StringBuilder htmlStyle) {
        String cssColor = getBackgroundColor(cellStyle);
        if (cssColor.length() > 0) {
            htmlStyle.append("background-color:").append(cssColor).append(";");
        }
    }

    protected abstract String getBackgroundColor(CellStyle cellStyle);
}
