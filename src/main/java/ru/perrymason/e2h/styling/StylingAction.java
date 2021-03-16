package ru.perrymason.e2h.styling;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import ru.perrymason.e2h.Excel2Html;

public interface StylingAction {
    void perform(Cell cell, CellStyle cellStyle, Excel2Html.CellSpans cellSpans, StringBuilder htmlStyle);
}
