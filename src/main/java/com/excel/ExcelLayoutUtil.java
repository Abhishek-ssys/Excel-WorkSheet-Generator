package com.excel;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Builds the sheet layout ONLY.
 * Parsing and fill logic are preserved (date loop, txt mapping, weekend rules).
 */
public class ExcelLayoutUtil {

    // === Title Rows ===
    public static void addTitleRows(Sheet sheet, int monthNum, int year, Map<String, CellStyle> styles) {
        // Title
        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(26);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("Performance Sheet - " + java.time.Month.of(monthNum).name() + " " + year);
        titleCell.setCellStyle(styles.get("title"));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 9));

        // Subtitle
        Row subtitleRow = sheet.createRow(1);
        subtitleRow.setHeightInPoints(20);
        Cell subTitleCell = subtitleRow.createCell(0);
        subTitleCell.setCellValue("Monthly Worksheet - " + java.time.Month.of(monthNum).name() + " " + year);
        subTitleCell.setCellStyle(styles.get("subtitle"));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 9));
    }

    // === Employee Info ===
    public static void addEmployeeInfo(Sheet sheet, Properties props, Map<String, CellStyle> styles) {
        String employeeName = props.getProperty("name", "");
        String projectName  = props.getProperty("projectName", "");
        String managerName  = props.getProperty("managerName", "");
        String employeeId   = props.getProperty("employeeId", "");

        // Labels (grey) + Values (white), borders everywhere
        Row infoRow1 = sheet.createRow(3);
        infoRow1.setHeightInPoints(18);
        infoRow1.createCell(0).setCellValue("Employee Name");
        infoRow1.getCell(0).setCellStyle(styles.get("label"));
        infoRow1.createCell(1).setCellValue(employeeName);
        infoRow1.getCell(1).setCellStyle(styles.get("value"));

        infoRow1.createCell(3).setCellValue("Project Name");
        infoRow1.getCell(3).setCellStyle(styles.get("label"));
        infoRow1.createCell(4).setCellValue(projectName);
        infoRow1.getCell(4).setCellStyle(styles.get("value"));

        Row infoRow2 = sheet.createRow(4);
        infoRow2.setHeightInPoints(18);
        infoRow2.createCell(0).setCellValue("Manager Name");
        infoRow2.getCell(0).setCellStyle(styles.get("label"));
        infoRow2.createCell(1).setCellValue(managerName);
        infoRow2.getCell(1).setCellStyle(styles.get("value"));

        infoRow2.createCell(3).setCellValue("Employee ID");
        infoRow2.getCell(3).setCellStyle(styles.get("label"));
        infoRow2.createCell(4).setCellValue(employeeId);
        infoRow2.getCell(4).setCellStyle(styles.get("value"));
    }

    // === Table Header ===
    public static void addTableHeader(Sheet sheet, Map<String, CellStyle> styles) {
        Row headerRow = sheet.createRow(6);
        headerRow.setHeightInPoints(20);

        headerRow.createCell(0).setCellValue("Date");
        headerRow.getCell(0).setCellStyle(styles.get("header"));

        headerRow.createCell(1).setCellValue("Task");
        headerRow.getCell(1).setCellStyle(styles.get("header"));
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 1, 9));
    }

    /**
     * === Fill Dates & Tasks ===
     * Keeps your core logic:
     * - Date iteration and key format
     * - Weekend detection (Sun, 2nd Sat, 4th Sat)
     * - Merge task columns 1..9 for each row
     * - Paste tasks from txt (joined with newline)
     * Only visual change: zebra striping for weekdays, pastel for weekend.
     */
    public static void fillDatesAndTasks(Sheet sheet, int monthNum, int year,
                                         Map<String, List<String>> dateTasks,
                                         Map<String, CellStyle> styles) {
        LocalDate startDate = LocalDate.of(year, monthNum, 1);
        LocalDate endDate = startDate.withDayOfMonth(startDate.lengthOfMonth());
        DateTimeFormatter fileDateFormatter = DateTimeFormatter.ofPattern("MMM_dd_yyyy", Locale.ENGLISH);

        int rowNum = 7;
        boolean zebra = false; // toggles only across weekdays

        for (LocalDate date = startDate; !date.isAfter(endDate); date = date.plusDays(1)) {
            Row row = sheet.createRow(rowNum++);
            row.setHeightInPoints(28);
            String dateKey = date.format(fileDateFormatter).toLowerCase();

            boolean isSunday     = date.getDayOfWeek() == DayOfWeek.SUNDAY;
            boolean isSecondSat  = date.getDayOfWeek() == DayOfWeek.SATURDAY && ((date.getDayOfMonth() - 1) / 7 + 1) == 2;
            boolean isFourthSat  = date.getDayOfWeek() == DayOfWeek.SATURDAY && ((date.getDayOfMonth() - 1) / 7 + 1) == 4;
            boolean isWeekendOff = (isSunday || isSecondSat || isFourthSat);

            // Date cell
            Cell dateCell = row.createCell(0);
            dateCell.setCellValue(dateKey);

            // Merge task cells 1..9 (unchanged)
            sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 1, 9));

            // Prepare the (single) visible task cell at col=1
            Cell taskCell = row.createCell(1);

            if (isWeekendOff) {
                // Weekend styling (pastel blue) + "Week Off"
                dateCell.setCellStyle(styles.get("weekendDate"));
                taskCell.setCellValue("Week Off");
                taskCell.setCellStyle(styles.get("weekendTask"));

                // Create the hidden merged cells 2..9 with same weekend border/fill for a clean block
                for (int col = 2; col <= 9; col++) {
                    Cell c = row.createCell(col);
                    c.setCellStyle(styles.get("weekendTask"));
                }
            } else {
                // Zebra striping for weekdays
                CellStyle dateStyle = zebra ? styles.get("date") : styles.get("dateAlt");
                CellStyle taskStyle = zebra ? styles.get("task") : styles.get("taskAlt");

                dateCell.setCellStyle(dateStyle);

                List<String> tasks = dateTasks.getOrDefault(dateKey, new ArrayList<>());
                taskCell.setCellValue(tasks.isEmpty() ? "" : String.join("\n", tasks));
                taskCell.setCellStyle(taskStyle);

                // Fill the merged companions 2..9 with same zebra style for seamless block
                for (int col = 2; col <= 9; col++) {
                    Cell c = row.createCell(col);
                    c.setCellStyle(taskStyle);
                }

                zebra = !zebra; // toggle only after a weekday row
            }
        }
    }
}
