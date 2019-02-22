package com.company;

public class Main {

    public static void main(String[] args) throws Exception {
        String fileSource = "ICT-maintenance_SVRHPSS_10.204.38.231.txt";
        String xlsxName = "ICT-maintenance_SVRHPSS_10.204.38.231.xlsx";

        TxtToXlsx export = new TxtToXlsx(fileSource, xlsxName);

        export.readFileInList();
        export.exportToXlsx();
    }
}

