package ru.perrymason.e2h.styling;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import ru.perrymason.e2h.Excel2Html;

public abstract class RotationStylingAction implements StylingAction {

    @Override
    public void perform(Cell cell, CellStyle cellStyle, Excel2Html.CellSpans cellSpans, StringBuilder htmlStyle) {
        short degree = getCssRotation(cellStyle.getRotation());
        if (degree != 0) {
            htmlStyle.append("transform:rotate(").append(degree).append("deg);");
            htmlStyle.append("-webkit-transform:rotate(").append(degree).append("deg);");
            htmlStyle.append("-ms-transform:rotate(").append(degree).append("deg);");
            htmlStyle.append("-moz-transform:rotate(").append(degree).append("deg);");
        }
    }

    protected abstract short getCssRotation(short rotation);
}
