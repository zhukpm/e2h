package ru.perrymason.e2h.styling;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import ru.perrymason.e2h.Excel2Html;

public abstract class BorderStylingAction extends BorderColorStylingAction {

    @Override
    public void perform(Cell cell, CellStyle cellStyle, Excel2Html.CellSpans cellSpans, StringBuilder htmlStyle) {
        htmlStyle.append(getBorderStyle(XSSFCellBorder.BorderSide.TOP, cellStyle.getBorderTopEnum(), cellStyle));
        htmlStyle.append(getBorderStyle(XSSFCellBorder.BorderSide.RIGHT, cellStyle.getBorderRightEnum(), cellStyle));
        htmlStyle.append(getBorderStyle(XSSFCellBorder.BorderSide.BOTTOM, cellStyle.getBorderBottomEnum(), cellStyle));
        htmlStyle.append(getBorderStyle(XSSFCellBorder.BorderSide.LEFT, cellStyle.getBorderLeftEnum(), cellStyle));
    }

    private String getBorderStyle(XSSFCellBorder.BorderSide border, BorderStyle style, CellStyle cellStyle) {
        String borderStyle = "";
        switch (style) {
            case NONE:
                return "";
            case HAIR:
            case THIN:
                borderStyle = "thin solid" + getBorderColor(border, cellStyle);
                break;
            case MEDIUM:
                borderStyle = "medium solid" + getBorderColor(border, cellStyle);
                break;
            case DASH_DOT_DOT:
            case DASH_DOT:
            case DASHED:
                borderStyle = "thin dashed" + getBorderColor(border, cellStyle);
                break;
            case DOTTED:
                borderStyle = "medium dotted" + getBorderColor(border, cellStyle);
                break;
            case THICK:
                borderStyle = "thick solid" + getBorderColor(border, cellStyle);
                break;
            case DOUBLE:
                borderStyle = "medium double" + getBorderColor(border, cellStyle);
                break;
            case SLANTED_DASH_DOT:
            case MEDIUM_DASH_DOT_DOT:
            case MEDIUM_DASH_DOT:
            case MEDIUM_DASHED:
                borderStyle = "medium dashed" + getBorderColor(border, cellStyle);
                break;
        }
        return "border-" + getCssBorder(border) + ":" + borderStyle + ";";
    }

    private String getCssBorder(XSSFCellBorder.BorderSide side) {
        switch (side) {
            case TOP:
                return "top";
            case BOTTOM:
                return "bottom";
            case LEFT:
                return "left";
            case RIGHT:
                return "right";
            default:
                return "";
        }
    }

}
