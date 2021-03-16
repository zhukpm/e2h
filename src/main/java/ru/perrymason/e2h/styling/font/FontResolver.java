package ru.perrymason.e2h.styling.font;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

public interface FontResolver {

    Font getFont(CellStyle cellStyle);

    String getDefaultFontFamilies();

}
