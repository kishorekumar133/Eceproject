package com.example.transformdemo;
//package com.example.transformdemo.service;

//public class TranscalcService {
//}
//package com.example.transformdemo.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeMap;

@Service
public class TranscalcService {
    public void generateExcel(int numSets) throws IOException {
        // Create workbook and sheet
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet("Student Data");

        // Create a row object
        XSSFRow row;

        // Scanner to get input (this should be modified to accept data from the web form)
        Scanner scanner = new Scanner(System.in);

        // Store the values in a map
        Map<String, Object[]> studentData = new TreeMap<>();
        for (int set = 1; set <= numSets; set++) {
            System.out.println("For Set " + set + ":");
            System.out.print("Enter KVA: ");
            double kva = scanner.nextDouble();
            System.out.print("Enter V: ");
            double v = scanner.nextDouble();

            // Calculate current (I)
            double i = (kva * 1000) / (1.732 * v);

            // Add data to the map
            studentData.put(String.valueOf(set), new Object[]{kva, v, i});
        }

        // Write the data into the sheet
        Set<String> keyid = studentData.keySet();
        int rowid = 0;
        for (String key : keyid) {
            row = spreadsheet.createRow(rowid++);
            Object[] objectArr = studentData.get(key);
            int cellid = 0;
            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue(obj.toString());
            }
        }

        // Write the workbook into the file
        FileOutputStream out = new FileOutputStream(new File("D:/Ece Project/transformdemo/GFGsheet.xlsx"));
        workbook.write(out);
        out.close();
        workbook.close();
    }
}

