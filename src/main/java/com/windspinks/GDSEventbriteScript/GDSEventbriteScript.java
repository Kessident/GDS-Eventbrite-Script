package com.windspinks.GDSEventbriteScript;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.csv.*;

import java.io.*;
import java.time.LocalDateTime;
import java.util.*;

public class GDSEventbriteScript {
    //private static final String SAMPLE_CSV_FILE_PATH = "/Users/windspinks/Google Drive/PRUMC Forms/Eventbrite Report Script/Sample CSV from Eventbrite.csv";
    private static final String EVENTBRITE_CSV_PATH = "Eventbrite CSV.csv";
    private static String[] columns = {"Vol #", "First Name", "Last Name", "Home Address 1", "HA 2", "City", "ST", "ZIP", "Cell Phone", "V?", "1?", "Age", "1ST", "2ND", "3RD"};


    public static void main(String[] args) throws IOException, InvalidFormatException{
        /*
        //Reading Excel File to find colors
        final String EXCEL_FILE_PATH = "Eventbrite template.xlsx";
        Workbook workbook = WorkbookFactory.create(new File(EXCEL_FILE_PATH));
        Sheet sheet = workbook.getSheetAt(0);

        Row second = sheet.getRow(1);

        int cellIterator = 1;
        for (Cell cell : second){
            CellStyle style = cell.getCellStyle();
            System.out.println("Cell " + cellIterator++);
            System.out.println("border top " + style.getBorderTop());
            System.out.println("border right " + style.getBorderRight());
            System.out.println("border left " + style.getBorderLeft());
            System.out.println("border bottom " + style.getBorderBottom() + "\n");
        }*/

        ArrayList<Volunteer> volunteerList = new ArrayList<>();
        Reader eventbriteCSVData = new FileReader(EVENTBRITE_CSV_PATH);
        Iterable<CSVRecord> volunteerCSVList = CSVFormat.EXCEL.withFirstRecordAsHeader().parse(eventbriteCSVData);
        for (CSVRecord volunteer : volunteerCSVList) {
            Volunteer currentVolunteer = new Volunteer();
            currentVolunteer.setFirstName(volunteer.get("First Name"));
            currentVolunteer.setLastName(volunteer.get("Last Name"));
            currentVolunteer.setHomeAddress(volunteer.get("Home Address 1"));
            currentVolunteer.setHomeAddressTwo(volunteer.get("Home Address 2"));
            currentVolunteer.setCity(volunteer.get("Home City"));
            currentVolunteer.setState(volunteer.get("Home State"));
            currentVolunteer.setZipCode(volunteer.get("Home Zip"));
            currentVolunteer.setCellPhone(volunteer.get("Cell Phone"));
            currentVolunteer.setVisitor(volunteer.get("Are you a Visitor to PRUMC?"));
            currentVolunteer.setFirstTime(volunteer.get("Is this your first Great Day of Service?"));
            currentVolunteer.setAge(volunteer.get("Age?"));
            currentVolunteer.setFirstChoice(volunteer.get(23));
            currentVolunteer.setSecondChoice(volunteer.get(24));
            currentVolunteer.setThirdChoice(volunteer.get(25));
            currentVolunteer.setEmailAddress(volunteer.get("Email"));
            currentVolunteer.setSpecial(volunteer.get(19));
            currentVolunteer.trimChoices();
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

        headerCellStyle.setBorderTop(BorderStyle.THIN);
        headerCellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        headerCellStyle.setBorderRight(BorderStyle.THIN);
        headerCellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        headerCellStyle.setBorderBottom(BorderStyle.THIN);
        headerCellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        headerCellStyle.setBorderLeft(BorderStyle.THIN);
        headerCellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());


        //Vol# Column Style
        CellStyle volColumnStyle = workbook.createCellStyle();
        volColumnStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        volColumnStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        volColumnStyle.setBorderTop(BorderStyle.THIN);
        volColumnStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        volColumnStyle.setBorderRight(BorderStyle.THIN);
        volColumnStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        volColumnStyle.setBorderBottom(BorderStyle.THIN);
        volColumnStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        volColumnStyle.setBorderLeft(BorderStyle.THIN);
        volColumnStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());


        //Volunteer Information Style
        CellStyle volunteerInfoStyle = workbook.createCellStyle();
        volunteerInfoStyle.setBorderTop(BorderStyle.THIN);
        volunteerInfoStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        volunteerInfoStyle.setBorderRight(BorderStyle.THIN);
        volunteerInfoStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        volunteerInfoStyle.setBorderBottom(BorderStyle.THIN);
        volunteerInfoStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        volunteerInfoStyle.setBorderLeft(BorderStyle.THIN);
        volunteerInfoStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());


        //Initial Row - Page num, date
        LocalDateTime localDateTime = LocalDateTime.now();
        String currentMonth = localDateTime.getMonth().name();
        int currentDay = localDateTime.getDayOfMonth();
        int currentYear = localDateTime.getYear();
        String currentDateFormatted = currentMonth + " " + currentDay + ", " + currentYear;

        Cell pageTitle = row.createCell(1);
        pageTitle.setCellValue("Eventbrite Signups " + currentDateFormatted);
        Cell pageNum = row.createCell(13);
        pageNum.setCellValue("P " + (rowNum / 32) + "/" + volunteerList.size() / 16);


        //Header Row
        Row headerRow = sheet.createRow(1);
        for (int i = 0; i < columns.length; i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        //Volunteer Rows
        int rowNum = 2;
        for (Volunteer vol : volunteerList){
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellStyle(volColumnStyle);

            Cell firstNameCell = row.createCell(1);
            firstNameCell.setCellValue(vol.getFirstName());
            firstNameCell.setCellStyle(volunteerInfoStyle);
            Cell lastNameCell = row.createCell(2);
            lastNameCell.setCellValue(vol.getLastName());
            lastNameCell.setCellStyle(volunteerInfoStyle);
            Cell homeAddressCell = row.createCell(3);
            homeAddressCell.setCellValue(vol.getHomeAddress());
            homeAddressCell.setCellStyle(volunteerInfoStyle);
            Cell homeAddress2Cell = row.createCell(4);
            homeAddress2Cell.setCellValue(vol.getHomeAddressTwo());
            homeAddress2Cell.setCellStyle(volunteerInfoStyle);
            Cell cityCell = row.createCell(5);
            cityCell.setCellValue(vol.getCity());
            cityCell.setCellStyle(volunteerInfoStyle);
            Cell stateCell = row.createCell(6);
            stateCell.setCellValue(vol.getState());
            stateCell.setCellStyle(volunteerInfoStyle);
            Cell zipCodeCell = row.createCell(7);
            zipCodeCell.setCellValue(vol.getZipCode());
            zipCodeCell.setCellStyle(volunteerInfoStyle);
            Cell cellPhoneCell = row.createCell(8);
            cellPhoneCell.setCellValue(vol.getCellPhone());
            cellPhoneCell.setCellStyle(volunteerInfoStyle);
            Cell isVisitorCell = row.createCell(9);
            isVisitorCell.setCellValue(vol.getVisitor());
            isVisitorCell.setCellStyle(volunteerInfoStyle);
            Cell isFirstTimeCell = row.createCell(10);
            isFirstTimeCell.setCellValue(vol.getFirstTime());
            isFirstTimeCell.setCellStyle(volunteerInfoStyle);
            Cell ageCell = row.createCell(11);
            ageCell.setCellValue(vol.getAge());
            ageCell.setCellStyle(volunteerInfoStyle);
            Cell firstChoiceCell = row.createCell(12);
            firstChoiceCell.setCellValue(vol.getFirstChoice());
            firstChoiceCell.setCellStyle(volunteerInfoStyle);
            Cell secondChoiceCell = row.createCell(13);
            secondChoiceCell.setCellValue(vol.getSecondChoice());
            secondChoiceCell.setCellStyle(volunteerInfoStyle);
            Cell thirdChoiceCell = row.createCell(14);
            thirdChoiceCell.setCellValue(vol.getThirdChoice());
            thirdChoiceCell.setCellStyle(volunteerInfoStyle);

            Row newRow = sheet.createRow(rowNum++);
            newRow.createCell(0).setCellStyle(volColumnStyle);

            Cell emailCell = newRow.createCell(2);
            emailCell.setCellValue(vol.getEmailAddress());
//            emailCell.setCellStyle(volunteerInfoStyle);
            Cell specialCell = newRow.createCell(4);
            specialCell.setCellValue(vol.getSpecial());
//            specialCell.setCellStyle(volunteerInfoStyle);
        }

        //Resize all Columns
//        for (int i = 0; i < columns.length; i++){
//            sheet.autoSizeColumn(i);
//        }
        sheet.setColumnWidth(0,40);
        sheet.setColumnWidth(1,100);
        sheet.setColumnWidth(2,120);
        sheet.setColumnWidth(3,200);
        sheet.setColumnWidth(4,70);
        sheet.setColumnWidth(5,85);
        sheet.setColumnWidth(6,25);
        sheet.setColumnWidth(7,40);
        sheet.setColumnWidth(8,85);
        sheet.setColumnWidth(9,25);
        sheet.setColumnWidth(10,25);
        sheet.setColumnWidth(11,30);
        sheet.setColumnWidth(12,30);
        sheet.setColumnWidth(13,30);
        sheet.setColumnWidth(14,30);


        //Write File
        FileOutputStream fileOut = new FileOutputStream("Formatted-Volunteer-List.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        workbook.close();
    }
}
