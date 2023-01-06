package com.example.demo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.util.Date;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

@RestController
@RequestMapping("/v1/demo")
public class DemoController {

    @GetMapping(value = "")
    public ResponseEntity<?> teste() {
        String str = "String teste";
        return new ResponseEntity<>(str, HttpStatus.OK);
    }

    @PostMapping(value = "/excel")
    public ResponseEntity<?> exportExcel() {

        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[]{1, "Amit", "Shukla"});
        data.put("2", new Object[]{2, "Lokesh", "Gupta"});
        data.put("3", new Object[]{3, "John", "Adwards"});
        data.put("4", new Object[]{4, "Brian", "Schultz"});

        String[] HEADERs = {"Id", "Name", "Lastname"};

        try (
                XSSFWorkbook workbook = new XSSFWorkbook();
                ByteArrayOutputStream out = new ByteArrayOutputStream();
        ) {

            XSSFSheet sheet = workbook.createSheet("Tutorials");

            // Header
            Row headerRow = sheet.createRow(0);
            for (int col = 0; col < HEADERs.length; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(HEADERs[col]);
            }

            int rowIdx = 1;

            Set<String> keyset = data.keySet();

            for (String key : keyset) {

                Row row = sheet.createRow(rowIdx++);

                Object[] objArr = data.get(key);

                int cellnum = 0;

                for (Object obj : objArr) {

                    Cell cell = row.createCell(cellnum++);

                    if (obj instanceof String)
                        cell.setCellValue((String) obj);
                    else if (obj instanceof Integer)
                        cell.setCellValue((Integer) obj);

                }

            }

            workbook.write(out);

            // Write the workbook in file system
            FileOutputStream outLocal = new FileOutputStream(new File("demo_" + new Date().getTime() + ".xlsx"));
            workbook.write(outLocal);
            outLocal.close();

            InputStreamResource file = new InputStreamResource(new ByteArrayInputStream(out.toByteArray()));

            String fileName = "demo_" + new Date().getTime() + ".xlsx";

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fileName)
                    .contentType(MediaType.parseMediaType(MediaType.APPLICATION_OCTET_STREAM_VALUE))
                    //.contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
                    .body(file);

        } catch (Exception e) {

            e.printStackTrace();

            return new ResponseEntity<String>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);

        }

    }

}
