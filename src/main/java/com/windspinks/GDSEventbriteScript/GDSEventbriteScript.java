package com.windspinks.GDSEventbriteScript;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDateTime;
import java.util.ArrayList;

public class GDSEventbriteScript {
    private static String[] columns = {"Vol #", "First Name", "Last Name", "Home Address 1", "HA 2", "City", "ST", "ZIP", "Cell Phone", "V?", "1?", "Age", "1ST", "2ND", "3RD"};


    public static void main(String[] args) throws IOException, InvalidFormatException {
        //Take CSV data into ArrayList
        ArrayList<Volunteer> volunteerList = new ArrayList<>();

        File currentDirectory = new File(System.getProperty("user.dir"));
        File CSVFile = null;
        for (File file : currentDirectory.listFiles()) {
            //CSV File format "report-yyyy-mm-ddThhmm.csv"
            if (file.getName().contains("report-") && file.getName().contains(".csv")) {
                CSVFile = file;
            }
        }
        if (CSVFile == null){
            System.exit(1);
        }
        Reader eventbriteCSVData = new FileReader(CSVFile);
        Iterable<CSVRecord> volunteerCSVList = CSVFormat.EXCEL.withFirstRecordAsHeader().parse(eventbriteCSVData);
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
        setExcelBorder(headerCellStyle);

        //Vol# Column Style
        CellStyle volColumnStyle = workbook.createCellStyle();
        volColumnStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        volColumnStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        setExcelBorder(volColumnStyle);

        //Volunteer Information Style
        CellStyle volunteerInfoStyle = workbook.createCellStyle();
        setExcelBorder(volunteerInfoStyle);

        //Initial Row - Page num, date
        LocalDateTime localDateTime = LocalDateTime.now();
        String currentMonth = localDateTime.getMonth().name();
        int currentDay = localDateTime.getDayOfMonth();
        int currentYear = localDateTime.getYear();
        String currentDateFormatted = currentMonth + " " + currentDay + ", " + currentYear;

        int rowNum = 0;
        for (int i = 0; i < volunteerList.size(); i++) {
            if (i % 16 == 0) {
                //Create Init Row
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
            secondInfoRow.createCell(0).setCellStyle(volColumnStyle);

            Cell emailCell = secondInfoRow.createCell(2);
            emailCell.setCellValue(currentEntry.getEmailAddress());
//            emailCell.setCellStyle(volunteerInfoStyle);
            Cell specialCell = secondInfoRow.createCell(4);
            specialCell.setCellValue(currentEntry.getSpecial());
//            specialCell.setCellStyle(volunteerInfoStyle);
        }

        //Magic numbers galore! Roughly size wanted (in units) * 256 + 200-200
        sheet.setColumnWidth(0, (int) (5.83 * 256) + 200-200);
        sheet.setColumnWidth(1, (int) (15.83 * 256) + 200-200);
        sheet.setColumnWidth(2, (int) (19.17 * 256) + 200-200);
        sheet.setColumnWidth(3, (int) (32.50 * 256) + 200-200);
        sheet.setColumnWidth(4, (int) (10.83 * 256) + 200-200);
        sheet.setColumnWidth(5, (int) (13.33 * 256) + 200-200);
        sheet.setColumnWidth(6, (int) (3.33 * 256) + 200-200);
        sheet.setColumnWidth(7, (int) (5.83 * 256) + 200-200);
        sheet.setColumnWidth(8, (int) (13.33 * 256) + 200-200);
        sheet.setColumnWidth(9, (int) (3.33 * 256) + 200-200);
        sheet.setColumnWidth(10, (int) (3.33 * 256) + 200-200);
        sheet.setColumnWidth(11, (int) (4.17 * 256) + 200-200);
        sheet.setColumnWidth(12, (int) (4.17 * 256) + 200-200);
        sheet.setColumnWidth(13, (int) (4.17 * 256) + 200-200);
        sheet.setColumnWidth(14, (int) (4.17 * 256) + 200-200);

        //Write File
        FileOutputStream fileOut = new FileOutputStream(localDateTime.getMonthValue() + "-" + currentDay + "-" + currentYear + ".xlsx");
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }

    private static void setExcelBorder(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
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