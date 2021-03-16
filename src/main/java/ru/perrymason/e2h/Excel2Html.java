package ru.perrymason.e2h;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import ru.perrymason.e2h.styling.*;
import ru.perrymason.e2h.styling.font.*;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;
import java.util.ArrayList;
import java.util.EnumSet;
import java.util.List;

/**
 * Converts HSSF- and XSSF- sheets to html tables
 */
public abstract class Excel2Html {

    private final Sheet workingSheet;
    private final CellRangeAddress range;
    private final List<CellRangeAddress> mergedCells;
    private final List<CellRangeAddress> closedMergedRegions;
    private final List<StylingAction> stylingAlgorithm;

    private final EnumSet<E2HOption> options = EnumSet.noneOf(E2HOption.class);

    private DataFormatter dataFormatter;
    private FormulaEvaluator formulaEvaluator;

    protected Excel2Html(Sheet sheet) {
        this.workingSheet = sheet;

        Row firstRow = sheet.getRow(sheet.getFirstRowNum());
        if (firstRow == null) {
            throw new IllegalArgumentException("Worksheet must contain at least 1 row");
        }
        this.range = new CellRangeAddress(sheet.getFirstRowNum(), sheet.getLastRowNum(),
                firstRow.getFirstCellNum(), firstRow.getLastCellNum());

        this.dataFormatter = new DataFormatter();

        this.mergedCells = new ArrayList<CellRangeAddress>();
        for (CellRangeAddress merged : sheet.getMergedRegions()) {
            if (this.range.intersects(merged)) {
                this.mergedCells.add(merged);
            }
        }
        this.closedMergedRegions = new ArrayList<CellRangeAddress>();

        this.stylingAlgorithm = new ArrayList<StylingAction>();
    }

    protected Excel2Html(Sheet sheet, CellRangeAddress range) {
        this.workingSheet = sheet;
        this.range = range;
        this.dataFormatter = new DataFormatter();

        this.range.validate(this.workingSheet.getWorkbook().getSpreadsheetVersion());

        this.mergedCells = new ArrayList<CellRangeAddress>();
        for (CellRangeAddress merged : sheet.getMergedRegions()) {
            if (this.range.intersects(merged)) {
                this.mergedCells.add(merged);
            }
        }
        this.closedMergedRegions = new ArrayList<CellRangeAddress>();

        this.stylingAlgorithm = new ArrayList<StylingAction>();
    }

    /**
     * Creates a new <tt>Excel2Html</tt> converter for a given sheet.
     * <p>Sheet can't be changed after creating a new converter. To convert multiple sheets you must create different
     * instances of this class.</p>
     * @param sheet a <tt>Sheet</tt> of HSSF- or XSSF- workbook
     * @return new instance of <tt>Excel2Html</tt> converter
     */
    public static Excel2Html getConverter(Sheet sheet) {
        Workbook workbook = sheet.getWorkbook();
        if (workbook instanceof XSSFWorkbook) {
            return new XSSF2Html(sheet);
        } else if (workbook instanceof HSSFWorkbook) {
            return new HSSF2Html(sheet);
        }
        throw new IllegalArgumentException("Workbook type must be either XSSF or HSSF");
    }

    /**
     * Creates a new <tt>Excel2Html</tt> converter for a given sheet. This converter will create an html table only for
     * cells specified in range.
     * <p>Sheet and range can't be changed after creating new converter. To convert multiple sheets you must create
     * different instances of this class.</p>
     * @param sheet a <tt>Sheet</tt> of HSSF- or XSSF- workbook
     * @param range a <tt>CellRangeAddress</tt> which this class will convert to html table (i.e. "C3:G18")
     * @return new instance of <tt>Excel2Html</tt> converter
     */
    public static Excel2Html getConverter(Sheet sheet, CellRangeAddress range) {
        Workbook workbook = sheet.getWorkbook();
        if (workbook instanceof XSSFWorkbook) {
            return new XSSF2Html(sheet, range);
        } else if (workbook instanceof HSSFWorkbook) {
            return new HSSF2Html(sheet, range);
        }
        throw new IllegalArgumentException("Workbook type must be either XSSF or HSSF");
    }

    public boolean addOption(E2HOption option) {
        return this.options.add(option);
    }

    public boolean addOption(EnumSet<E2HOption> options) {
        return this.options.addAll(options);
    }

    public boolean removeOption(E2HOption option) {
        return this.options.remove(option);
    }

    public boolean removeOption(EnumSet<E2HOption> options) {
        return this.options.removeAll(options);
    }

    public boolean hasOption(E2HOption option) {
        return this.options.contains(option);
    }

    public boolean hasOption(EnumSet<E2HOption> options) {
        return this.options.containsAll(options);
    }

    public EnumSet<E2HOption> getOptions() {
        return EnumSet.copyOf(this.options);
    }

    public void replaceOptions(EnumSet<E2HOption> options) {
        this.options.clear();
        this.options.addAll(options);
    }

    public DataFormatter getDataFormatter() {
        return dataFormatter;
    }

    public void setDataFormatter(DataFormatter dataFormatter) {
        this.dataFormatter = dataFormatter;
    }

    /**
     * Writes an html table to the specified <tt>OutputStream</tt>
     * <p>Note that this method uses <tt>UTF-8</tt> encoding for characters</p>
     * @param outputStream
     * @throws XMLStreamException
     */
    public void writeHtml(OutputStream outputStream) throws XMLStreamException {
        try {
            writeHtml(new OutputStreamWriter(outputStream, "utf-8"));
        } catch (UnsupportedEncodingException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Writes an html table using the specified <tt>Writer</tt>
     * @param writer
     * @throws XMLStreamException
     */
    public void writeHtml(Writer writer) throws XMLStreamException {
        buildStylingAlgorithm();
        XMLStreamWriter out = XMLOutputFactory.newInstance().createXMLStreamWriter(writer);
//        writeHtmlHeaders(out);
        out.writeStartElement("table");
        out.writeAttribute("style", "border-collapse: collapse;");
        int firstRow = range.getFirstRow();
        if (hasOption(E2HOption.USE_TABLE_HEADERS)) {
            firstRow += writeTableHeader(out);
        }
        for (int rowNum = firstRow; rowNum <= range.getLastRow(); rowNum++) {
            Row row = workingSheet.getRow(rowNum);
            if (row == null) {
                // Write an empty row with default height
                out.writeStartElement("tr");
                out.writeAttribute("style", "height:15pt;");
                out.writeEndElement();
                continue;
            }
            out.writeStartElement("tr");
            if (hasOption(E2HOption.CELL_HEIGHT)) {
                out.writeAttribute("style", "height:" + row.getHeightInPoints() + "pt;");
            }
            for (int cellNum = range.getFirstColumn(); cellNum <= range.getLastColumn(); cellNum++) {
                Cell cell = row.getCell(cellNum);
                if (cell == null) {
                    // write an empty cell
                    out.writeStartElement("td");
                    out.writeEndElement();
                    continue;
                }
                if (isSpanned(cell)) {
                    continue;
                }
                out.writeStartElement("td");
                writeCell(cell, out);
                out.writeEndElement();
            }
            out.writeEndElement();
        }
        out.writeEndElement();
        out.writeEndDocument();
        out.close();
    }

    private void buildStylingAlgorithm() {
        final boolean evaluateFormulas = hasOption(E2HOption.EVALUATE_FORMULAS);
        if (evaluateFormulas) {
            formulaEvaluator = workingSheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
        }

        if (hasOption(E2HOption.HORIZONTAL_ALIGNMENT)) {
            stylingAlgorithm.add(new HorizontalAlignmentStylingAction() {
                @Override
                protected boolean isEvaluateFormulas() {
                    return evaluateFormulas;
                }
            });
        }

        if (hasOption(E2HOption.VERTICAL_ALIGNMENT)) {
            stylingAlgorithm.add(new VerticalAlignmentStylingAction());
        }

        if (hasOption(E2HOption.BORDER_STYLE)) {
            if (hasOption(E2HOption.BORDER_COLOR)) {
                stylingAlgorithm.add(getBorderStylingAction());
            } else {
                stylingAlgorithm.add(new BorderStylingAction() {
                    @Override
                    protected String getBorderColor(XSSFCellBorder.BorderSide border, CellStyle cellStyle) {
                        return "";
                    }
                });
            }
        } else if (hasOption(E2HOption.BORDER_COLOR)) {
            // Border-colors without borders
            stylingAlgorithm.add(getBorderColorOnlyStylingAction());
        }

        if (hasOption(E2HOption.CELL_WIDTH)) {
            stylingAlgorithm.add(new WidthStylingAction());
        }

        if (hasOption(E2HOption.CELL_BACKGROUND_COLOR)) {
            stylingAlgorithm.add(getBackgroundColorStylingAction());
        }

        if (hasOption(EnumSet.of(E2HOption.FONT_SIZE, E2HOption.FONT_FAMILY))) {
            // We can wrap styles up
            if (hasOption(E2HOption.FONT_STYLE)) {
                stylingAlgorithm.add(new FullBundledFontStylingAction(getFontResolver()));
            } else {
                stylingAlgorithm.add(new BundledFontStylingAction(getFontResolver()));
            }
        } else {
            // We should add different styles for every value
            if (hasOption(E2HOption.FONT_FAMILY)) {
                stylingAlgorithm.add(new FontFamilyStylingAction(getFontResolver()));
            }

            if (hasOption(E2HOption.FONT_SIZE)) {
                stylingAlgorithm.add(new FontSizeStylingAction(getFontResolver()));
            }

            if (hasOption(E2HOption.FONT_STYLE)) {
                stylingAlgorithm.add(new FontStyleStylingAction(getFontResolver()));
            }
        }

        if (hasOption(E2HOption.FONT_COLOR)) {
            stylingAlgorithm.add(getFontColorStylingAction());
        }

        if (hasOption(E2HOption.TEXT_ROTATION)) {
            stylingAlgorithm.add(getRotationStylingAction());
        }
    }

    protected abstract String getCssColor(Color color);

    protected abstract StylingAction getBorderStylingAction();

    protected abstract StylingAction getBorderColorOnlyStylingAction();

    protected abstract StylingAction getBackgroundColorStylingAction();

    protected abstract FontResolver getFontResolver();

    protected abstract StylingAction getFontColorStylingAction();

    protected abstract StylingAction getRotationStylingAction();

//    private void writeHtmlHeaders(XMLStreamWriter out) throws XMLStreamException {
//        out.writeStartElement("html");
//
//        out.writeStartElement("head");
//        out.writeStartElement("meta");
//        out.writeAttribute("charset", "UTF-8");
//        out.writeEndElement();
//        out.writeEndElement();
//    }

    private int writeTableHeader(XMLStreamWriter out) throws XMLStreamException {
        int headerRows = 1;
        for (int i = 0; i < headerRows; i++) {
            out.writeStartElement("tr");
            Row row = workingSheet.getRow(range.getFirstRow() + i);
            if (hasOption(E2HOption.CELL_HEIGHT)) {
                out.writeAttribute("style", "height:" + row.getHeightInPoints() + "pt;");
            }
            for (Cell cell : row) {
                if (isSpanned(cell)) {
                    continue;
                }
                out.writeStartElement("th");
                CellSpans spans = writeCell(cell, out);
                if (spans != null) {
                    headerRows = Math.max(headerRows, spans.rowspan);
                }
                out.writeEndElement();
            }
            out.writeEndElement();
        }
        return headerRows;
    }

    private CellSpans writeCell(Cell cell, XMLStreamWriter out) throws XMLStreamException {
        CellSpans cellSpans = getCellSpans(cell);
        if (cellSpans != null) {
            if (cellSpans.colspan > 1) {
                out.writeAttribute("colspan", String.valueOf(cellSpans.colspan));
            }
            if (cellSpans.rowspan > 1) {
                out.writeAttribute("rowspan", String.valueOf(cellSpans.rowspan));
            }
        }
        encodeCellStyle(cell, cellSpans, out);
        encodeCellValue(cell, out);
        return cellSpans;
    }

    private boolean isSpanned(Cell cell) {
        for (CellRangeAddress rangeAddress : closedMergedRegions) {
            if (rangeAddress.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                return true;
            }
        }
        return false;
    }

    private void encodeCellStyle(Cell cell, CellSpans cellSpans, XMLStreamWriter out) throws XMLStreamException {
        CellStyle cellStyle = cell.getCellStyle();
        StringBuilder htmlStyle = new StringBuilder();
        for (StylingAction action : stylingAlgorithm) {
            action.perform(cell, cellStyle, cellSpans, htmlStyle);
        }
        if (htmlStyle.length() > 0) {
            out.writeAttribute("style", htmlStyle.toString());
        }
    }

    private CellSpans getCellSpans(Cell cell) {
        if (mergedCells.size() == 0) {
            return null;
        }
        CellRangeAddress merged = null;
        for (CellRangeAddress rangeAddress : mergedCells) {
            if (rangeAddress.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                merged = rangeAddress;
                break;
            }
        }
        if (merged != null) {
            mergedCells.remove(merged);
            closedMergedRegions.add(merged);
            CellSpans spans = new CellSpans();
            spans.rowspan = Math.min(merged.getLastRow(), range.getLastRow()) - cell.getRowIndex() + 1;
            spans.colspan = Math.min(merged.getLastColumn(), range.getLastColumn()) - cell.getColumnIndex() + 1;
            return spans;
        }
        return null;
    }

    private void encodeCellValue(Cell cell, XMLStreamWriter out) throws XMLStreamException {
        out.writeCharacters(dataFormatter.formatCellValue(cell, formulaEvaluator));
    }

    public class CellSpans {
        int colspan;
        int rowspan;

        public int getColspan() {
            return colspan;
        }

        public int getRowspan() {
            return rowspan;
        }
    }
}
