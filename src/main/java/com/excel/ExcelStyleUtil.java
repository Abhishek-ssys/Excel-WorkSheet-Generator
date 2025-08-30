package com.excel;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * Centralized style factory (no template file needed).
 * Matches the provided design (Kaushik.xlsx):
 * - Title/subtitle bars
 * - Dark header with white text
 * - Zebra weekdays (white / light grey)
 * - Pastel weekend
 * - Borders everywhere
 * - Calibri font
 */
public class ExcelStyleUtil {

    // Core palette (RGB)
    private static final short[] BLUE_DARK   = new short[]{  0,  51, 102};  // header/title
    private static final short[] GREY_MED    = new short[]{191, 191, 191};  // subtitle / label bg
    private static final short[] GREY_LIGHT  = new short[]{242, 242, 242};  // zebra alt
    private static final short[] WHITE       = new short[]{255, 255, 255};  // base
    private static final short[] WEEKEND_BG  = new short[]{221, 235, 247};  // pastel blue weekend

    public static Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<>();

        // Title & Subtitle
        styles.put("title",   make(wb, BLUE_DARK,  true, 16, HorizontalAlignment.CENTER, IndexedColors.WHITE.getIndex(),  true, false));
        styles.put("subtitle",make(wb, GREY_MED,   true, 12, HorizontalAlignment.CENTER, IndexedColors.BLACK.getIndex(),  true, false));

        // Labels (Employee Name, etc.) + Values
        styles.put("label",   make(wb, GREY_MED,   true, 11, HorizontalAlignment.LEFT,   IndexedColors.BLACK.getIndex(),  true, false));
        styles.put("value",   make(wb, WHITE,     false, 11, HorizontalAlignment.LEFT,   IndexedColors.BLACK.getIndex(),  true, false));

        // Table Headers
        styles.put("header",  make(wb, BLUE_DARK,  true, 11, HorizontalAlignment.CENTER, IndexedColors.WHITE.getIndex(),  true, false));

        // Weekday cells (zebra striping)
        styles.put("date",     make(wb, WHITE,      false, 11, HorizontalAlignment.CENTER, IndexedColors.BLACK.getIndex(), true, false));
        styles.put("dateAlt",  make(wb, GREY_LIGHT, false, 11, HorizontalAlignment.CENTER, IndexedColors.BLACK.getIndex(), true, false));
        styles.put("task",     make(wb, WHITE,     false, 11, HorizontalAlignment.LEFT,   IndexedColors.BLACK.getIndex(), true, true));
        styles.put("taskAlt",  make(wb, GREY_LIGHT,false, 11, HorizontalAlignment.LEFT,   IndexedColors.BLACK.getIndex(), true, true));

        // Weekend cells
        styles.put("weekendDate", make(wb, WEEKEND_BG, true, 11, HorizontalAlignment.CENTER, IndexedColors.BLACK.getIndex(), true, false));
        styles.put("weekendTask", make(wb, WEEKEND_BG, false, 11, HorizontalAlignment.LEFT,   IndexedColors.BLACK.getIndex(), true, true));

        return styles;
    }

    private static CellStyle make(Workbook wb,
                                  short[] bgRgb,
                                  boolean bold,
                                  int fontPt,
                                  HorizontalAlignment align,
                                  short fontColorIdx,
                                  boolean withBorders,
                                  boolean wrap) {
        CellStyle cs = wb.createCellStyle();

        // Font
        Font f = wb.createFont();
        f.setFontName("Calibri");
        f.setBold(bold);
        f.setFontHeightInPoints((short) fontPt);
        f.setColor(fontColorIdx);
        cs.setFont(f);

        // Background fill
        ((XSSFCellStyle) cs).setFillForegroundColor(new XSSFColor(
                new java.awt.Color(bgRgb[0], bgRgb[1], bgRgb[2]), null));
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Alignment & wrapping
        cs.setAlignment(align);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        cs.setWrapText(wrap);

        // Borders
        if (withBorders) {
            cs.setBorderBottom(BorderStyle.THIN);
            cs.setBorderTop(BorderStyle.THIN);
            cs.setBorderLeft(BorderStyle.THIN);
            cs.setBorderRight(BorderStyle.THIN);
        }
        return cs;
    }
}
