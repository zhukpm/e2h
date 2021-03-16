package ru.perrymason.e2h;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import ru.perrymason.e2h.styling.*;
import ru.perrymason.e2h.styling.font.FontResolver;
import ru.perrymason.e2h.styling.font.FontStylingAction;

/**
 * Implements HSSF-specific options such as getting colors, fonts, rotation, etc.
 */
class HSSF2Html extends Excel2Html implements FontResolver {

    private final HSSFWorkbook workbook;
    private final HSSFPalette palette;

    HSSF2Html(Sheet sheet) {
        super(sheet);
        workbook = (HSSFWorkbook) sheet.getWorkbook();
        palette = workbook.getCustomPalette();
    }

    HSSF2Html(Sheet sheet, CellRangeAddress range) {
        super(sheet, range);
        workbook = (HSSFWorkbook) sheet.getWorkbook();
        palette = workbook.getCustomPalette();
    }

    @Override
    protected String getCssColor(Color color) {
        // TODO
        if (color == null) {
            return "black";
        }
        HSSFColor hssfColor = (HSSFColor) color;
        return "rgb(" + hssfColor.getTriplet()[0] + "," + hssfColor.getTriplet()[1] + "," + hssfColor.getTriplet()[2] + ")";
    }

    @Override
    protected StylingAction getBorderStylingAction() {
        return new BorderStylingAction() {
            @Override
            protected String getBorderColor(XSSFCellBorder.BorderSide border, CellStyle cellStyle) {
                switch (border) {
                    case TOP:
                        return " " + getCssColor(palette.getColor(cellStyle.getTopBorderColor()));
                    case RIGHT:
                        return " " + getCssColor(palette.getColor(cellStyle.getRightBorderColor()));
                    case BOTTOM:
                        return " " + getCssColor(palette.getColor(cellStyle.getBottomBorderColor()));
                    case LEFT:
                        return " " + getCssColor(palette.getColor(cellStyle.getLeftBorderColor()));
                }
                return "";
            }
        };
    }

    @Override
    protected StylingAction getBorderColorOnlyStylingAction() {
        return new BorderColorOnlyStylingAction() {
            @Override
            protected String getBorderColor(XSSFCellBorder.BorderSide border, CellStyle cellStyle) {
                switch (border) {
                    case TOP:
                        return getCssColor(palette.getColor(cellStyle.getTopBorderColor()));
                    case RIGHT:
                        return getCssColor(palette.getColor(cellStyle.getRightBorderColor()));
                    case BOTTOM:
                        return getCssColor(palette.getColor(cellStyle.getBottomBorderColor()));
                    case LEFT:
                        return getCssColor(palette.getColor(cellStyle.getLeftBorderColor()));
                }
                return "";
            }
        };
    }

    @Override
    protected StylingAction getBackgroundColorStylingAction() {
        return new BackgroundColorStylingAction() {
            @Override
            protected String getBackgroundColor(CellStyle cellStyle) {
                HSSFCellStyle style = (HSSFCellStyle) cellStyle;
                HSSFColor color = palette.getColor(style.getFillForegroundColor());
                if (color == null) {
                    color = palette.getColor(style.getFillBackgroundColor());
                }
                if (color != null && !color.equals(HSSFColor.AUTOMATIC.getInstance())) {
                    return getCssColor(color);
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
                HSSFFont font = (HSSFFont) getFont(cellStyle);
                String color = getCssColor(font.getHSSFColor(workbook));
                htmlStyle.append("color:").append(color).append(";");
            }
        };
    }

    @Override
    protected StylingAction getRotationStylingAction() {
        return new RotationStylingAction() {
            @Override
            protected short getCssRotation(short rotation) {
                return (short) -rotation;
            }
        };
    }

    @Override
    public Font getFont(CellStyle cellStyle) {
        return ((HSSFCellStyle) cellStyle).getFont(workbook);
    }

    @Override
    public String getDefaultFontFamilies() {
        return HSSFFont.FONT_ARIAL + ",sans-serif";
    }
}
