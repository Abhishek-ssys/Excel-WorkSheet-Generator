package com.excel;

import java.io.*;
import java.nio.file.Files;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.Month;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

public class MonthExcelGenerator {

    // === RGB Color Config ===
    private static final short[] SUB_HEADER_COLOR = new short[]{255, 204, 153}; // light orange
    private static final short[] HEADER_COLOR = new short[]{204, 229, 255}; // light blue
    private static final short[] LABEL_COLOR = new short[]{224, 224, 224}; // light grey
    private static final short[] WEEKEND_COLOR = new short[]{255, 230, 230}; // light red
    private static final short[] NORMAL_COLOR = new short[]{255, 255, 255}; // white



    public static void main(String[] args) throws Exception {
        Properties props = loadInputFile();
        int monthNum = Integer.parseInt(props.getProperty("month", "1"));
        int year = Integer.parseInt(props.getProperty("year", String.valueOf(LocalDate.now().getYear())));

        Workbook workbook = createWorkbook(monthNum, year,props);
        String fileName = "Monthly WorkSheet-" + Month.of(monthNum).name().toLowerCase() + "_" + year + ".xlsx";
        saveWorkbook(workbook, fileName);

        System.out.println("Abhishek generated A Worksheet: " + fileName);

        // === Trigger Jira Extraction ===
        extractAndWriteJiras(new File("."));
    }




    // === Workbook Creator ===
    private static Workbook createWorkbook(int monthNum, int year, Properties props) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(String.valueOf(Month.of(monthNum)) + year);

        // Styles
        CellStyle headerStyle = createRGBStyle(workbook, HEADER_COLOR, true, HorizontalAlignment.CENTER, "Book Antiqua");
        CellStyle subHeaderStyle = createRGBStyle(workbook, SUB_HEADER_COLOR, true, HorizontalAlignment.CENTER, "Calibri");
        CellStyle labelStyle = createRGBStyle(workbook, LABEL_COLOR, true, HorizontalAlignment.LEFT, "Calibri");
        CellStyle weekendStyle = createRGBStyle(workbook, WEEKEND_COLOR, false, HorizontalAlignment.LEFT, "Calibri");
        CellStyle weekendTaskStyle = createRGBStyle(workbook, WEEKEND_COLOR, false, HorizontalAlignment.CENTER, "Calibri");
        CellStyle dateBorderStyle = createBorderedStyle(workbook, NORMAL_COLOR, true, HorizontalAlignment.LEFT, "Calibri");
        CellStyle taskBorderStyle = createBorderedStyle(workbook, NORMAL_COLOR, false, HorizontalAlignment.LEFT, "Calibri");

        // Sections
        addTitleRows(sheet, monthNum, year, headerStyle, subHeaderStyle);
        addEmployeeInfo(sheet, labelStyle,props);
        addTableHeader(sheet, headerStyle);

        Map<String, List<String>> dateTasks = loadTasksFromTxt();
        fillDatesAndTasks(sheet, monthNum, year, dateTasks, dateBorderStyle, taskBorderStyle, weekendStyle, weekendTaskStyle);

        // Auto-size
        for (int i = 0; i <= 6; i++) sheet.autoSizeColumn(i);

        return workbook;
    }

    // === Title Rows ===
    private static void addTitleRows(Sheet sheet, int monthNum, int year, CellStyle headerStyle, CellStyle subHeaderStyle) {
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("Performance Sheet - " + Month.of(monthNum).name() + " " + year);
        titleCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 9));

        Row subtitleRow = sheet.createRow(1);
        Cell subTitleCell = subtitleRow.createCell(0);
        subTitleCell.setCellValue("Monthly Worksheet - " + Month.of(monthNum).name() + " " + year);
        subTitleCell.setCellStyle(subHeaderStyle);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 9));
    }

    // === Employee Info (Dynamic from input.txt instead of config.properties) ===
    private static void addEmployeeInfo(Sheet sheet, CellStyle labelStyle, Properties props) {

        String employeeName = props.getProperty("name", "");
        String projectName = props.getProperty("projectName", "");
        String managerName = props.getProperty("managerName", "");
        String employeeId = props.getProperty("employeeId", "");

        Row infoRow1 = sheet.createRow(3);
        infoRow1.createCell(0).setCellValue("Employee Name");
        infoRow1.getCell(0).setCellStyle(labelStyle);
        infoRow1.createCell(1).setCellValue(employeeName);

        infoRow1.createCell(3).setCellValue("Project Name");
        infoRow1.getCell(3).setCellStyle(labelStyle);
        infoRow1.createCell(4).setCellValue(projectName);

        Row infoRow2 = sheet.createRow(4);
        infoRow2.createCell(0).setCellValue("Manager Name");
        infoRow2.getCell(0).setCellStyle(labelStyle);
        infoRow2.createCell(1).setCellValue(managerName);

        infoRow2.createCell(3).setCellValue("Employee ID");
        infoRow2.getCell(3).setCellStyle(labelStyle);
        infoRow2.createCell(4).setCellValue(employeeId);
    }


    // === Table Header ===
    private static void addTableHeader(Sheet sheet, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(6);
        headerRow.createCell(0).setCellValue("Date");
        headerRow.getCell(0).setCellStyle(headerStyle);
        headerRow.createCell(1).setCellValue("Task");
        headerRow.getCell(1).setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 1, 9));
    }

    // === Fill Dates & Tasks ===
    private static void fillDatesAndTasks(Sheet sheet, int monthNum, int year,
                                          Map<String, List<String>> dateTasks,
                                          CellStyle dateBorderStyle, CellStyle taskBorderStyle,
                                          CellStyle weekendStyle, CellStyle weekendTaskStyle) {
        LocalDate startDate = LocalDate.of(year, monthNum, 1);
        LocalDate endDate = startDate.withDayOfMonth(startDate.lengthOfMonth());
        DateTimeFormatter fileDateFormatter = DateTimeFormatter.ofPattern("MMM_dd_yyyy", Locale.ENGLISH);

        int rowNum = 7;
        for (LocalDate date = startDate; !date.isAfter(endDate); date = date.plusDays(1)) {
            Row row = sheet.createRow(rowNum++);
            String dateKey = date.format(fileDateFormatter).toLowerCase();

            // Date cell
            Cell dateCell = row.createCell(0);
            dateCell.setCellValue(dateKey);
            dateCell.setCellStyle(dateBorderStyle);

            // Task merged cells
            sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 1, 9));
            for (int col = 1; col <= 9; col++) {
                Cell taskCell = row.createCell(col);
                taskCell.setCellStyle(taskBorderStyle);

                if (col == 1) {
                    boolean isSunday = date.getDayOfWeek() == DayOfWeek.SUNDAY;
                    boolean isSecondSat = date.getDayOfWeek() == DayOfWeek.SATURDAY && ((date.getDayOfMonth() - 1) / 7 + 1) == 2;
                    boolean isFourthSat = date.getDayOfWeek() == DayOfWeek.SATURDAY && ((date.getDayOfMonth() - 1) / 7 + 1) == 4;

                    if (isSunday || isSecondSat || isFourthSat) {
                        taskCell.setCellValue("Week Off");
                        taskCell.setCellStyle(weekendTaskStyle);
                        dateCell.setCellStyle(weekendStyle);
                    } else {
                        List<String> tasks = dateTasks.getOrDefault(dateKey, new ArrayList<>());
                        taskCell.setCellValue(tasks.isEmpty() ? "" : String.join("\n", tasks));
                    }
                }
            }
        }
    }

    // === Save Workbook ===
    private static void saveWorkbook(Workbook workbook, String fileName) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(fileName)) {
            workbook.write(fos);
        }
        workbook.close();
    }

    // === Styles ===
    private static CellStyle createRGBStyle(Workbook wb, short[] rgb, boolean bold, HorizontalAlignment align, String fontName) {
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBold(bold);
        font.setFontName(fontName);
        style.setFont(font);
        ((XSSFCellStyle) style).setFillForegroundColor(new XSSFColor(new java.awt.Color(rgb[0], rgb[1], rgb[2]), null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(align);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static CellStyle createBorderedStyle(Workbook wb, short[] rgb, boolean bold, HorizontalAlignment align, String fontName) {
        return createRGBStyle(wb, rgb, bold, align, fontName);
    }

    // === Load Tasks from TXT ===
    private static Map<String, List<String>> loadTasksFromTxt() throws IOException {
        Map<String, List<String>> dateTasks = new HashMap<>();
        File folder = new File(".");
        File[] files = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".txt"));

        if (files != null) {
            for (File file : files) {
                String fileName = file.getName().replace(".txt", "").toLowerCase();
                List<String> lines = Files.readAllLines(file.toPath());
                lines.removeIf(s -> s == null || s.trim().isEmpty());
                dateTasks.computeIfAbsent(fileName, k -> new ArrayList<>()).addAll(lines);
            }
        }
        return dateTasks;
    }

    // === Extract Jira Tickets ===
    private static void extractAndWriteJiras(File folder) throws IOException {
        File[] files = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".txt"));
        if (files == null) return;

        // Allowed Jira project keys
        Set<String> allowedProjects = new HashSet<>();
        allowedProjects.add("HDAG");
        allowedProjects.add("HCAG");
        allowedProjects.add("HCCUG");
        allowedProjects.add("HDCUG");
        allowedProjects.add("APIGW");
        allowedProjects.add("ESB");
        allowedProjects.add("KAFKA");
        allowedProjects.add("OHAB");

        // Regex: jira + space + (PROJECT DIGITS)
        Pattern jiraPattern = Pattern.compile("(?i)\\bjira\\s+([a-zA-Z]+)[- ]?(\\d+)\\b");
        Set<String> jiraTickets = new HashSet<>();

        for (File file : files) {
            List<String> lines = Files.readAllLines(file.toPath());
            for (String line : lines) {
                Matcher matcher = jiraPattern.matcher(line);
                while (matcher.find()) {
                    String project = matcher.group(1).toUpperCase();  // e.g. "hdag" -> "HDAG"
                    String number = matcher.group(2);                 // digits part

                    if (allowedProjects.contains(project)) {
                        String ticket = project + "-" + number;       // normalize: PROJECT-123
                        jiraTickets.add(ticket);
                    }
                }
            }
        }

        // Write all unique Jiras to a file
        File output = new File(folder, "All_Jiras.txt");
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(output))) {
            for (String ticket : jiraTickets) {
                writer.write(ticket);
                writer.newLine();
            }
        }

        System.out.println("Extracted " + jiraTickets.size() + " Jira tickets into " + output.getAbsolutePath());
    }
    private static Properties loadInputFile() {
        Properties props = new Properties();
        File inputFile = new File("input.txt");
        if (inputFile.exists()) {
            try (FileInputStream fis = new FileInputStream(inputFile)) {
                props.load(fis);
                System.out.println("Loaded input.txt successfully.");
            } catch (IOException e) {
                System.err.println("Failed to load input.txt: " + e.getMessage());
            }
        } else {
            System.err.println("No input.txt found! Using defaults.");
        }
        return props;
    }


}
