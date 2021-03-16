package ru.perrymason.e2h.styling;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import ru.perrymason.e2h.Excel2Html;

public class WidthStylingAction implements StylingAction {
    @Override
    public void perform(Cell cell, CellStyle cellStyle, Excel2Html.CellSpans cellSpans, StringBuilder htmlStyle) {
        if (cellSpans == null || cellSpans.getColspan() == 1) {
            htmlStyle.append("width:").append(cell.getSheet().getColumnWidthInPixels(cell.getColumnIndex())).append("px;");
        } else {
            int lastColShift = cellSpans.getColspan() - 1;
            float width = 0;
            for (int col = 0; col < lastColShift; col++) {
                width += cell.getSheet().getColumnWidthInPixels(cell.getColumnIndex() + col);
            }
            htmlStyle.append("width:").append(width).append("px;");
        }
    }
}
