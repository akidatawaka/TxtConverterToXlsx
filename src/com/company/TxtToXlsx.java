package com.company;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

public class TxtToXlsx {
    private String fileName;
    private List<String> list;

    public TxtToXlsx(String _fileName) {
        this.fileName = _fileName;
    }

    protected void readFileInList() {
        List<String> lines = Collections.emptyList();
        try {
            lines = Files.readAllLines(Paths.get(fileName), StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
        }
        list = lines;
    }

    protected void exportToXlsx(){
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Export");

        int rowNum = 0;


        Iterator<String> iterator = list.iterator();
        while (iterator.hasNext()) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;

            String[] _column = iterator.next().toString().split(";");
            for(int j = 0 ; j < _column.length ; j++)
            {
                Cell cell = row.createCell(colNum++);
                cell.setCellValue(_column[j]);
            }
        }
        try
        {
            FileOutputStream outputStream = new FileOutputStream("Export.xlsx");
            workbook.write(outputStream);
            workbook.close();
        } catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}
