package com.windspinks.GDSEventbriteScript;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.csv.*;

import java.io.*;
import java.util.*;

public class GDSEventbriteScript {
    private static final String SAMPLE_CSV_FILE_PATH = "/Users/windspinks/Google Drive/PRUMC Forms/Eventbrite Report Script/Sample CSV from Eventbrite.csv";
    private static String[] columns = {"Vol #", "First Name", "Last Name", "Home Address 1", "HA 2", "City", "ST", "ZIP", "Cell Phone", "V?", "1?", "Age", "1ST", "2ND", "3RD"};


    public static void main(String[] args) throws IOException, InvalidFormatException{
        ArrayList<Volunteer> volunteerList = new ArrayList<Volunteer>();

        Reader eventbriteCSVData = new FileReader(SAMPLE_CSV_FILE_PATH);
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

        //Header Font
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
//        headerFont.setColor(IndexedColors.RED.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setFillBackgroundColor(IndexedColors.BLUE.getIndex());

        Row headerRow = sheet.createRow(1);

        //Initial Row - Page num, date


        //Header Row
        for (int i = 0; i < columns.length; i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        //Volunteer Rows
        int rowNum = 2;
        for (Volunteer vol : volunteerList){
            Row row = sheet.createRow(rowNum++);
            row.createCell(1).setCellValue(vol.getFirstName());
            row.createCell(2).setCellValue(vol.getLastName());
            row.createCell(3).setCellValue(vol.getHomeAddress());
            row.createCell(4).setCellValue(vol.getHomeAddressTwo());
            row.createCell(5).setCellValue(vol.getCity());
            row.createCell(6).setCellValue(vol.getState());
            row.createCell(7).setCellValue(vol.getZipCode());
            row.createCell(8).setCellValue(vol.getCellPhone());
            row.createCell(9).setCellValue(vol.getVisitor());
            row.createCell(10).setCellValue(vol.getFirstTime());
            row.createCell(11).setCellValue(vol.getAge());
            row.createCell(12).setCellValue(vol.getFirstChoice());
            row.createCell(13).setCellValue(vol.getSecondChoice());
            row.createCell(14).setCellValue(vol.getThirdChoice());
            Row newRow = sheet.createRow(rowNum++);
            newRow.createCell(2).setCellValue(vol.getEmailAddress());
            newRow.createCell(4).setCellValue(vol.getSpecial());
        }

        //Resize all Columns
        for (int i = 0; i < columns.length; i++){
            sheet.autoSizeColumn(i);
        }


        //Write File
        FileOutputStream fileOut = new FileOutputStream("Formatted-Volunteer-List.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        workbook.close();

    }
}
