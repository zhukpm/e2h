package ru.perrymason.e2h.styling;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import ru.perrymason.e2h.Excel2Html;

public class VerticalAlignmentStylingAction implements StylingAction {
    @Override
    public void perform(Cell cell, CellStyle cellStyle, Excel2Html.CellSpans cellSpans, StringBuilder htmlStyle) {
        switch (cellStyle.getVerticalAlignmentEnum()) {
            case TOP:
                htmlStyle.append("vertical-align:top;");
                break;
            case CENTER:
                htmlStyle.append("vertical-align:middle;");
                break;
            case BOTTOM:
            default:
                htmlStyle.append("vertical-align:bottom;");
        }
    }
}
