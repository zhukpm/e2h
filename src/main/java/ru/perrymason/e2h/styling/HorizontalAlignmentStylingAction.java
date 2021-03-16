package ru.perrymason.e2h.styling;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import ru.perrymason.e2h.Excel2Html;

public abstract class HorizontalAlignmentStylingAction implements StylingAction {

    @Override
    public void perform(Cell cell, CellStyle cellStyle, Excel2Html.CellSpans cellSpans, StringBuilder htmlStyle) {
        switch (cellStyle.getAlignmentEnum()) {
            case LEFT:
                htmlStyle.append("text-align:left;");
                break;
            case RIGHT:
                htmlStyle.append("text-align:right;");
                break;
            case CENTER:
                htmlStyle.append("text-align:center;");
                break;
            case JUSTIFY:
                htmlStyle.append("text-align:justify;");
                break;
            case GENERAL:
            default:
                switch (cell.getCellTypeEnum()) {
                    case NUMERIC:
                        htmlStyle.append("text-align:right;");
                        break;
                    case BOOLEAN:
                        htmlStyle.append("text-align:center;");
                        break;
                    case FORMULA:
                        if (isEvaluateFormulas()) {
                            htmlStyle.append("text-align:right;");
                            break;
                        }
                    case STRING:
                    default:
                        htmlStyle.append("text-align:left;");
                }
        }
    }

    protected abstract boolean isEvaluateFormulas();
}
