package com.company;

public class Main {

    public static void main(String[] args) throws Exception {
        String fileLocation = "ICT-maintenance_SVRHPSS_10.204.38.231.txt";
        TxtToXlsx export = new TxtToXlsx(fileLocation);

        export.readFileInList();
        export.exportToXlsx();
    }
}

