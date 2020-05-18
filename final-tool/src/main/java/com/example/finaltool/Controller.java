package com.example.finaltool;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

@org.springframework.stereotype.Controller
public class Controller {

    @PostMapping
    public ResponseEntity<InputStreamResource> importFile(@RequestParam("filename") MultipartFile reapExcelDataFile) throws IOException {
        List<String> phoneNums = new ArrayList<>();
        XSSFWorkbook workbook = new XSSFWorkbook(reapExcelDataFile.getInputStream());
        XSSFSheet worksheet = workbook.getSheetAt(0);

        for (int i = 0; i < worksheet.getPhysicalNumberOfRows(); i++) {
            XSSFRow row = worksheet.getRow(i);
            //String phoneNum=Double.toString(row.getCell(3).getNumericCellValue());
            String phoneNum = row.getCell(3).getStringCellValue();
            phoneNums.add(phoneNum);
        }

        for (int i = 0; i < phoneNums.size(); i++) {
            if (phoneNums.get(i).matches("^84\\d{9}$"))
                phoneNums.set(i, phoneNums.get(i).replaceFirst("84", "0"));
        }

        for (int i = 0; i < phoneNums.size(); i++) {
            if (phoneNums.get(i).matches("^0162\\d{6}$"))
                phoneNums.set(i, phoneNums.get(i).replaceFirst("0162", "032"));
            if (phoneNums.get(i).matches("^0163\\d{6}$"))
                phoneNums.set(i, phoneNums.get(i).replaceFirst("0163", "033"));
            if (phoneNums.get(i).matches("^0164\\d{6}$"))
                phoneNums.set(i, phoneNums.get(i).replaceFirst("0164", "043"));
            if (phoneNums.get(i).matches("^0165\\d{6}$"))
                phoneNums.set(i, phoneNums.get(i).replaceFirst("0165", "035"));
            if (phoneNums.get(i).matches("^0166\\d{6}$"))
                phoneNums.set(i, phoneNums.get(i).replaceFirst("0166", "036"));
            if (phoneNums.get(i).matches("^0167\\d{6}$"))
                phoneNums.set(i, phoneNums.get(i).replaceFirst("0167", "037"));
            if (phoneNums.get(i).matches("^0168\\d{6}$"))
                phoneNums.set(i, phoneNums.get(i).replaceFirst("0168", "038"));
            if (phoneNums.get(i).matches("^0169\\d{6}$"))
                phoneNums.set(i, phoneNums.get(i).replaceFirst("0169", "039"));
        }

        Workbook workbookObj=new XSSFWorkbook();
        ByteArrayOutputStream out =new ByteArrayOutputStream();
        Sheet sheet =workbookObj.createSheet();


        for (int i = 0; i < phoneNums.size(); i++) {
            Row row = sheet.createRow(i);
            row.createCell(0).setCellValue(phoneNums.get(i));
        }

        workbookObj.write(out);
        ByteArrayInputStream byteArrayInputStream= new ByteArrayInputStream(out.toByteArray());
        HttpHeaders headers= new HttpHeaders();
        headers.add("Content-Disposition", "attachment; filename=export.xlsx");
        return ResponseEntity
                .ok()
                .headers(headers)
                .body(new InputStreamResource(byteArrayInputStream));
    }
}

