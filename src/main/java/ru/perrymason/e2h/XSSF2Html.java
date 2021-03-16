package ru.perrymason.e2h;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import ru.perrymason.e2h.styling.*;
import ru.perrymason.e2h.styling.font.FontResolver;
import ru.perrymason.e2h.styling.font.FontStylingAction;

/**
 * Implements XSSF-specific options such as getting colors, fonts, rotation, etc.
 */
class XSSF2Html extends Excel2Html implements FontResolver {

    XSSF2Html(Sheet sheet) {
        super(sheet);
    }

    XSSF2Html(Sheet sheet, CellRangeAddress range) {
        super(sheet, range);
    }

    @Override
    protected String getCssColor(Color color) {
        if (color == null) {
            return "black";
        }

        String strColor = ((XSSFColor) color).getARGBHex();
        if (strColor == null) {
            return "black";
        }
        return "#" + strColor.substring(2);
    }

    @Override
    protected StylingAction getBorderStylingAction() {
        return new BorderStylingAction() {
            @Override
            protected String getBorderColor(XSSFCellBorder.BorderSide border, CellStyle cellStyle) {
                return " " + getCssColor(((XSSFCellStyle) cellStyle).getBorderColor(border));
            }
        };
    }

    @Override
    protected StylingAction getBorderColorOnlyStylingAction() {
        return new BorderColorOnlyStylingAction() {
            @Override
            protected String getBorderColor(XSSFCellBorder.BorderSide border, CellStyle cellStyle) {
                return getCssColor(((XSSFCellStyle) cellStyle).getBorderColor(border));
            }
        };
    }

    @Override
    protected StylingAction getBackgroundColorStylingAction() {
        return new BackgroundColorStylingAction() {
            @Override
            protected String getBackgroundColor(CellStyle cellStyle) {
                XSSFCellStyle style = (XSSFCellStyle) cellStyle;
                XSSFColor color = style.getFillBackgroundColorColor();
                if (color != null) {
                    String strColor = color.getARGBHex();
                    if (strColor != null) {
                        return "#" + strColor.substring(2);
                    } else {
                        color = style.getFillForegroundXSSFColor();
                        return getCssColor(color);
                    }
                }
                return "";
            }
        };
    }

    @Override
    protected FontResolver getFontResolver() {
        return this;
    }

    @Override
    protected StylingAction getFontColorStylingAction() {
        return new FontStylingAction(this) {
            @Override
            public void perform(Cell cell, CellStyle cellStyle, CellSpans cellSpans, StringBuilder htmlStyle) {
                XSSFFont font = (XSSFFont) getFont(cellStyle);
                String color = getCssColor(font.getXSSFColor());
                htmlStyle.append("color:").append(color).append(";");
            }
        };
    }

    @Override
    protected StylingAction getRotationStylingAction() {
        return new RotationStylingAction() {
            @Override
            protected short getCssRotation(short rotation) {
                if (rotation > 90) {
                    return (short) (rotation - 90);
                }
                return (short) -rotation;
            }
        };
    }

    @Override
    public Font getFont(CellStyle cellStyle) {
        return ((XSSFCellStyle) cellStyle).getFont();
    }

    @Override
    public String getDefaultFontFamilies() {
        return XSSFFont.DEFAULT_FONT_NAME + ",sans-serif";
    }
}
