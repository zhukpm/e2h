package ru.perrymason.e2h.styling.font;

import org.apache.poi.ss.usermodel.Font;
import ru.perrymason.e2h.styling.StylingAction;

public abstract class FontStylingAction implements StylingAction {

    protected final FontResolver fontResolver;

    protected FontStylingAction(FontResolver fontResolver) {
        this.fontResolver = fontResolver;
    }

    protected final String getFontFamilies(Font font) {
        String fontName = font.getFontName();
        if (fontName.length() == 0) {
            return fontResolver.getDefaultFontFamilies();
        } else {
            return fontName + "," + fontResolver.getDefaultFontFamilies();
        }
    }

    protected final String getBold(Font font) {
        return font.getBold() ? "bold " : "";
    }

    protected final String getItalic(Font font) {
        return font.getItalic() ? "italic " : "";
    }
}
