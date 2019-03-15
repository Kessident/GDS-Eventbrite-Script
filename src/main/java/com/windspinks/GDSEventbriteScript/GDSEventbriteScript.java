package com.windspinks.GDSEventbriteScript;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVRecord;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;

public class GDSEventbriteScript {
    private static String[] columns = {"Vol #", "First Name", "Last Name", "Home Address 1", "HA 2", "City", "ST", "ZIP", "Cell Phone", "V?", "1?", "Age", "1ST", "2ND", "3RD"};
    private static Logger logger = LogManager.getLogger(GDSEventbriteScript.class);

    public static void main(String[] args) {
        //Take CSV data into ArrayList
        ArrayList<Volunteer> volunteerList = new ArrayList<>();

        File currentDirectory = new File(System.getProperty("user.dir"));
        File CSVFile = null;
        File[] currDirectoryFiles = currentDirectory.listFiles();

        if (currDirectoryFiles == null) {
            logger.error("'user.dir' not a directory");
            System.exit(1);
        } else {
            Arrays.sort(currDirectoryFiles);
            for (File file : currDirectoryFiles) {
                //CSV File format "report-yyyy-mm-ddThhmm.csv"
                if (file.getName().contains("report-") && file.getName().contains(".csv")) {
                    CSVFile = file;
                }
            }
        }
<<<<<<< HEAD
        if (CSVFile == null) {
=======

        if (CSVFile == null) {
            logger.error("CSV File Not Found");
            System.exit(1);
        }

        Reader eventbriteCSVData = null;
        try {
            eventbriteCSVData = new FileReader(CSVFile);
        } catch (FileNotFoundException | NullPointerException ex) {
            logger.catching(ex);
            System.exit(1);
        }

        Iterable<CSVRecord> volunteerCSVList = null;
        try {
            volunteerCSVList = CSVFormat.EXCEL.withFirstRecordAsHeader().parse(eventbriteCSVData);
        } catch (IOException ex) {
            logger.catching(ex);
>>>>>>> dev
            System.exit(1);
        }

        for (CSVRecord CSVVolunteer : volunteerCSVList) {
            Volunteer currentVolunteer = new Volunteer();
            setVolunteerInfo(currentVolunteer, CSVVolunteer);
            volunteerList.add(currentVolunteer);
        }

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Volunteers");

        //Header Font and Style
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        setExcelBorder(headerCellStyle, true, true, true, true);

        //Vol# Column Style
        CellStyle volColumnStyle = workbook.createCellStyle();
        volColumnStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        volColumnStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        setExcelBorder(volColumnStyle, true, true, true, true);

        //Volunteer Information Style
        CellStyle volunteerInfoStyle = workbook.createCellStyle();
        setExcelBorder(volunteerInfoStyle, true, true, true, true);

        CellStyle rightBorderStyle = workbook.createCellStyle();
        setExcelBorder(rightBorderStyle, false, true, true, false);
        CellStyle bottomBorderStyle = workbook.createCellStyle();
        setExcelBorder(bottomBorderStyle, false, false, true, false);

        //Initial Row - Page num, date
        LocalDateTime localDateTime = LocalDateTime.now();
        String currentMonth = localDateTime.getMonth().name();
        int currentDay = localDateTime.getDayOfMonth();
        int currentYear = localDateTime.getYear();
        String currentDateFormatted = currentMonth + " " + currentDay + ", " + currentYear;

        int rowNum = 0;
        for (int i = 0; i < volunteerList.size(); i++) {
            if (i % 16 == 0) {
                //Create Inittial Row: pg Num, date
                Row initRow = sheet.createRow(rowNum++);
                Cell pageTitle = initRow.createCell(1);
                pageTitle.setCellValue("Eventbrite Signups " + currentDateFormatted);
                Cell pageNum = initRow.createCell(13);
                pageNum.setCellValue("P " + (rowNum / 32 + 1) + "/" + (volunteerList.size() / 16 + 1));

                //Header Row
                Row headerRow = sheet.createRow(rowNum++);
                for (int j = 0; j < columns.length; j++) {
                    Cell cell = headerRow.createCell(j);
                    cell.setCellValue(columns[j]);
                    cell.setCellStyle(headerCellStyle);
                }
            }

            //16 Vol Rows
            Volunteer currentEntry = volunteerList.get(i);
            Row firstInfoRow = sheet.createRow(rowNum++);
            firstInfoRow.createCell(0).setCellStyle(volColumnStyle);

            Cell firstNameCell = firstInfoRow.createCell(1);
            firstNameCell.setCellValue(currentEntry.getFirstName());
            firstNameCell.setCellStyle(volunteerInfoStyle);
            Cell lastNameCell = firstInfoRow.createCell(2);
            lastNameCell.setCellValue(currentEntry.getLastName());
            lastNameCell.setCellStyle(volunteerInfoStyle);
            Cell homeAddressCell = firstInfoRow.createCell(3);
            homeAddressCell.setCellValue(currentEntry.getHomeAddress());
            homeAddressCell.setCellStyle(volunteerInfoStyle);
            Cell homeAddress2Cell = firstInfoRow.createCell(4);
            homeAddress2Cell.setCellValue(currentEntry.getHomeAddressTwo());
            homeAddress2Cell.setCellStyle(volunteerInfoStyle);
            Cell cityCell = firstInfoRow.createCell(5);
            cityCell.setCellValue(currentEntry.getCity());
            cityCell.setCellStyle(volunteerInfoStyle);
            Cell stateCell = firstInfoRow.createCell(6);
            stateCell.setCellValue(currentEntry.getState());
            stateCell.setCellStyle(volunteerInfoStyle);
            Cell zipCodeCell = firstInfoRow.createCell(7);
            zipCodeCell.setCellValue(currentEntry.getZipCode());
            zipCodeCell.setCellStyle(volunteerInfoStyle);
            Cell cellPhoneCell = firstInfoRow.createCell(8);
            cellPhoneCell.setCellValue(currentEntry.getCellPhone());
            cellPhoneCell.setCellStyle(volunteerInfoStyle);
            Cell isVisitorCell = firstInfoRow.createCell(9);
            isVisitorCell.setCellValue(currentEntry.getVisitor());
            isVisitorCell.setCellStyle(volunteerInfoStyle);
            Cell isFirstTimeCell = firstInfoRow.createCell(10);
            isFirstTimeCell.setCellValue(currentEntry.getFirstTime());
            isFirstTimeCell.setCellStyle(volunteerInfoStyle);
            Cell ageCell = firstInfoRow.createCell(11);
            ageCell.setCellValue(currentEntry.getAge());
            ageCell.setCellStyle(volunteerInfoStyle);
            Cell firstChoiceCell = firstInfoRow.createCell(12);
            firstChoiceCell.setCellValue(currentEntry.getFirstChoice());
            firstChoiceCell.setCellStyle(volunteerInfoStyle);
            Cell secondChoiceCell = firstInfoRow.createCell(13);
            secondChoiceCell.setCellValue(currentEntry.getSecondChoice());
            secondChoiceCell.setCellStyle(volunteerInfoStyle);
            Cell thirdChoiceCell = firstInfoRow.createCell(14);
            thirdChoiceCell.setCellValue(currentEntry.getThirdChoice());
            thirdChoiceCell.setCellStyle(volunteerInfoStyle);


            Row secondInfoRow = sheet.createRow(rowNum++);
            for (int j = 1; j < 14; j++){
                secondInfoRow.createCell(j).setCellStyle(bottomBorderStyle);
            }

            secondInfoRow.createCell(0).setCellStyle(volColumnStyle);
            Cell emailCell = secondInfoRow.createCell(2);
            emailCell.setCellValue(currentEntry.getEmailAddress());
            emailCell.setCellStyle(bottomBorderStyle);
            Cell specialCell = secondInfoRow.createCell(4);
            specialCell.setCellValue(currentEntry.getSpecial());
            specialCell.setCellStyle(bottomBorderStyle);

            secondInfoRow.createCell(14).setCellStyle(rightBorderStyle);
        }

        //Magic numbers galore! Obtained from Template
        sheet.setColumnWidth(0, 1462);
        sheet.setColumnWidth(1, 3657);
        sheet.setColumnWidth(2, 4388);
        sheet.setColumnWidth(3, 7314);
        sheet.setColumnWidth(4, 2560);
        sheet.setColumnWidth(5, 3108);
        sheet.setColumnWidth(6, 914);
        sheet.setColumnWidth(7, 1462);
        sheet.setColumnWidth(8, 3108);
        sheet.setColumnWidth(9, 914);
        sheet.setColumnWidth(10, 914);
        sheet.setColumnWidth(11, 1097);
        sheet.setColumnWidth(12, 1097);
        sheet.setColumnWidth(13, 1097);
        sheet.setColumnWidth(14, 1097);

        //Margins
        sheet.setMargin(Sheet.TopMargin, 0.75);
        sheet.setMargin(Sheet.RightMargin, 0.25);
        sheet.setMargin(Sheet.BottomMargin, 0.75);
        sheet.setMargin(Sheet.LeftMargin, 0.25);

        sheet.getPrintSetup().setLandscape(true);
        sheet.getPrintSetup().setPaperSize(XSSFPrintSetup.LETTER_PAPERSIZE);
        //Write File
        FileOutputStream fileOut = null;
        try {
            // TODO: 3/14/19 Set filename as name of input file, not just current date? If two in same day, need different names, use hour?
            String currentYearFileName = currentDirectory + "/" + currentYear;

            if (Files.exists(Paths.get(currentYearFileName))) {
                fileOut = new FileOutputStream(currentYear + "/" + currentYear + "-" + localDateTime.getMonthValue() + "-" + currentDay + ".xlsx");
            } else {
                fileOut = new FileOutputStream(currentYear + "-" + localDateTime.getMonthValue() + "-" + currentDay + ".xlsx");
            }
        } catch (FileNotFoundException ex) {
            logger.catching(ex);
            System.exit(1);
        }

        try {
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
        } catch (IOException ex) {
            logger.catching(ex);
            System.exit(1);
        }
    }

    private static void setExcelBorder(CellStyle style, boolean doTop, boolean doRight, boolean doBottom, boolean doLeft) {
        if (doTop) {
            style.setBorderTop(BorderStyle.THIN);
            style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        }
        if (doRight) {
            style.setBorderRight(BorderStyle.THIN);
            style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        }
        if (doBottom) {
            style.setBorderBottom(BorderStyle.THIN);
            style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        }
        if (doLeft) {
            style.setBorderLeft(BorderStyle.THIN);
            style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        }
    }

    private static void setVolunteerInfo(Volunteer vol, CSVRecord CSVVolunteer) {
        vol.setFirstName(CSVVolunteer.get("First Name"));
        vol.setLastName(CSVVolunteer.get("Last Name"));
        vol.setHomeAddress(CSVVolunteer.get("Home Address 1"));
        vol.setHomeAddressTwo(CSVVolunteer.get("Home Address 2"));
        vol.setCity(CSVVolunteer.get("Home City"));
        vol.setState(CSVVolunteer.get("Home State"));
        vol.setZipCode(CSVVolunteer.get("Home Zip"));
        vol.setCellPhone(CSVVolunteer.get("Cell Phone"));
        vol.setVisitor(CSVVolunteer.get("Are you a Visitor to PRUMC?"));
        vol.setFirstTime(CSVVolunteer.get("Is this your first Great Day of Service?"));
        vol.setAge(CSVVolunteer.get("Age?"));
        vol.setFirstChoice(CSVVolunteer.get("First Choice of Project (38 Projects - be sure to scroll down!)"));
        vol.setSecondChoice(CSVVolunteer.get("Second Choice of Project (38 Projects - be sure to scroll down!)"));
        vol.setThirdChoice(CSVVolunteer.get("Third Choice of Project (38 Projects - be sure to scroll down!)"));
        vol.setEmailAddress(CSVVolunteer.get("Email"));
        vol.setSpecial(CSVVolunteer.get("Do you have any special skills, tools, equipment, etc?"));
        vol.trimChoices();
    }
}
