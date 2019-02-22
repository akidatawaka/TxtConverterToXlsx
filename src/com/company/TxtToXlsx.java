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
    private String exportFileName;
    private List<String> list;

    public TxtToXlsx(String fileName, String exportFileName) {
        this.fileName = fileName;
        this.exportFileName = exportFileName;
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
            FileOutputStream outputStream = new FileOutputStream(exportFileName);
            workbook.write(outputStream);
            workbook.close();
        } catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}
