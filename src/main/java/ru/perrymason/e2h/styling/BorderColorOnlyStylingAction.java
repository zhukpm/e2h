package ru.perrymason.e2h.styling;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import ru.perrymason.e2h.Excel2Html;

public abstract class BorderColorOnlyStylingAction extends BorderColorStylingAction {

    @Override
    public void perform(Cell cell, CellStyle cellStyle, Excel2Html.CellSpans cellSpans, StringBuilder htmlStyle) {
        htmlStyle.append("border-top-color:").append(getBorderColor(XSSFCellBorder.BorderSide.TOP, cellStyle));
        htmlStyle.append("border-right-color:").append(getBorderColor(XSSFCellBorder.BorderSide.RIGHT, cellStyle));
        htmlStyle.append("border-bottom-color:").append(getBorderColor(XSSFCellBorder.BorderSide.BOTTOM, cellStyle));
        htmlStyle.append("border-left-color:").append(getBorderColor(XSSFCellBorder.BorderSide.LEFT, cellStyle));
    }

}
