package ru.perrymason.e2h.styling;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

abstract class BorderColorStylingAction implements StylingAction {
    protected abstract String getBorderColor(XSSFCellBorder.BorderSide border, CellStyle cellStyle);
}
