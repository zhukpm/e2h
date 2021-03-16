package ru.perrymason.e2h;

import java.util.EnumSet;

/**
 * Options for {@link Excel2Html} converter
 */
public enum E2HOption {

    /**
     * If added to the converter, it will use {@code <th>} tags instead of {@code <td>} tags for the first row of the
     * html table. If one or more cells of the first row have <tt>rowspan &gt; 1</tt> then {@code <th>} tag will be used
     * for the first {@code max{rowspan}} rows.
     */
    USE_TABLE_HEADERS,

    /**
     * If added to the converter, it will try to evaluate formulas in an Excel sheet. If not, formulas will be displayed
     * as strings.
     */
    EVALUATE_FORMULAS,

    /**
     * If added to the converter, it will specify cell width as it is in an Excel sheet
     */
    CELL_WIDTH,
    /**
     * If added to the converter, it will specify cell height as it is in an Excel sheet
     */
    CELL_HEIGHT,

    /**
     * If added to the converter, it will specify font size as it is in an Excel sheet
     */
    FONT_SIZE,
    /**
     * If added to the converter, it will specify font family as it is in an Excel sheet. Additional families also will
     * be added for compatibility.
     */
    FONT_FAMILY,
    /**
     * If added to the converter, it will specify different font styles as it is in an Excel sheet.
     * <p>Currently these styles are supported: bold, italic, underlined</p>
     */
    FONT_STYLE,
    /**
     * If added to the converter, it will specify text color as it is in an Excel sheet.
     */
    FONT_COLOR,

    /**
     * If added to the converter, it will specify border colors as it is in an Excel sheet.
     */
    BORDER_COLOR,
    /**
     * If added to the converter, it will specify border types (and width) as it is in an Excel sheet.
     */
    BORDER_STYLE,

    /**
     * If added to the converter, it will specify background colors for cells as it is in an Excel sheet.
     */
    CELL_BACKGROUND_COLOR,

    /**
     * If added to the converter, it will specify horizontal alignment as it is in an Excel sheet.
     */
    HORIZONTAL_ALIGNMENT,
    /**
     * If added to the converter, it will specify vertical alignment as it is in an Excel sheet.
     */
    VERTICAL_ALIGNMENT,

    /**
     * If added to the converter, it will specify text-rotation as it is in an Excel sheet.
     */
    TEXT_ROTATION,;

    /**
     * If added to the converter, it will use all font options: {@link #FONT_SIZE}, {@link #FONT_STYLE},
     * {@link #FONT_FAMILY}, {@link #FONT_COLOR}.
     */
    public static final EnumSet<E2HOption> FONT = EnumSet.of(FONT_SIZE, FONT_STYLE, FONT_FAMILY, FONT_COLOR);
    /**
     * If added to the converter, it will use cell sizes options: {@link #CELL_WIDTH}, {@link #CELL_HEIGHT}.
     */
    public static final EnumSet<E2HOption> CELL_SIZES = EnumSet.of(CELL_HEIGHT, CELL_WIDTH);
    /**
     * If added to the converter, it will use borders options: {@link #BORDER_COLOR}, {@link #BORDER_STYLE}.
     */
    public static final EnumSet<E2HOption> BORDERS = EnumSet.of(BORDER_STYLE, BORDER_COLOR);
    /**
     * If added to the converter, it will use all colors options: {@link #CELL_BACKGROUND_COLOR}, {@link #FONT_COLOR},
     * {@link #BORDER_COLOR}.
     */
    public static final EnumSet<E2HOption> COLORS = EnumSet.of(CELL_BACKGROUND_COLOR, FONT_COLOR, BORDER_COLOR);
    /**
     * If added to the converter, it will use alignment options: {@link #VERTICAL_ALIGNMENT},
     * {@link #HORIZONTAL_ALIGNMENT}.
     */
    public static final EnumSet<E2HOption> ALIGNMENT = EnumSet.of(VERTICAL_ALIGNMENT, HORIZONTAL_ALIGNMENT);

    /**
     * If added to the converter, it will use all {@link E2HOption} options, except for {@link #USE_TABLE_HEADERS}.
     */
    public static final EnumSet<E2HOption> STANDARD_OPTIONS = EnumSet.range(EVALUATE_FORMULAS, TEXT_ROTATION);
}
