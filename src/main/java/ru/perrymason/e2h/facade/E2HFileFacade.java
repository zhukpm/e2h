package ru.perrymason.e2h.facade;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import ru.perrymason.e2h.E2HOption;
import ru.perrymason.e2h.Excel2Html;

import javax.xml.stream.XMLStreamException;
import java.io.*;
import java.util.EnumSet;

/**
 * File facade for {@link Excel2Html} converter.
 */
public final class E2HFileFacade implements Closeable {
    private final Workbook workbook;
    private final EnumSet<E2HOption> options = EnumSet.noneOf(E2HOption.class);

    private CellRangeAddress range;
    private Sheet workingSheet;

    /**
     * Creates new {@link Excel2Html} converter for a given Excel file.
     * <p>Note that {@code E2HFileFacade} should be closed after use in order to properly release resources.</p>
     * @param excel Excel file. Must be either 97-2003 (.xls) or 2007-* (.xlsx) file
     * @throws IOException if an error occurs while reading the data
     * @throws InvalidFormatException if the contents of the file cannot be parsed into a {@link Workbook}
     * @throws EncryptedDocumentException If the workbook given is password protected
     */
    public E2HFileFacade(File excel) throws IOException, InvalidFormatException {
        workbook = WorkbookFactory.create(excel);
        workingSheet = workbook.getSheetAt(0);
    }

    /**
     * Selects sheet to be converted to an html table, by its name
     * @param name sheet name
     * @throws IllegalArgumentException if there is no sheet with such name
     */
    public void selectSheet(String name) {
        Sheet temp = workbook.getSheet(name);
        if (temp == null) {
            throw new IllegalArgumentException("Specified sheet '" + name + "' doesn't exist in workbook");
        }
        workingSheet = temp;
    }

    /**
     * Selects sheet to be converted to an html table, by its index (0-based)
     * @param index
     */
    public void selectSheetAt(int index) {
        workingSheet = workbook.getSheetAt(index);
    }

    /**
     * Selects a cell range to be converted to an html table from a cell range reference string.
     * @param range a standard area ref (e.g. "B1:D8").
     */
    public void selectCellRange(String range) {
        this.range = CellRangeAddress.valueOf(range);
        this.range.validate(workbook.getSpreadsheetVersion());
    }

    /**
     * Selects a cell range to be converted to an html table from a cell range indexes. Indexes are 0-based
     * @param firstRow Index of first row
     * @param lastRow Index of last row (inclusive), must be equal to or larger than {@code firstRow}
     * @param firstCol Index of first column
     * @param lastCol Index of last column (inclusive), must be equal to or larger than {@code firstCol}
     */
    public void selectCellRange(int firstRow, int lastRow, int firstCol, int lastCol) {
        this.range = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        this.range.validate(workbook.getSpreadsheetVersion());
    }

    /**
     * Removes current cell range restrictions so the whole sheet will be converted to an html table.
     */
    public void selectCellRangeAsWholeSheet() {
        this.range = null;
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

    /**
     * Writes an html table to a given file.
     * @param file
     * @throws FileNotFoundException
     * @throws XMLStreamException
     */
    public void writeHtml(File file) throws FileNotFoundException, XMLStreamException {
        Excel2Html excel2Html;
        if (range == null) {
            excel2Html = Excel2Html.getConverter(workingSheet);
        } else {
            excel2Html = Excel2Html.getConverter(workingSheet, range);
        }
        excel2Html.replaceOptions(options);
        excel2Html.writeHtml(new FileOutputStream(file));
    }

    @Override
    public void close() throws IOException {
        workbook.close();
    }
}
