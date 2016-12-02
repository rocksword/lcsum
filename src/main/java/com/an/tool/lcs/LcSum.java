package com.an.tool.lcs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LcSum {
    private static final String LC_DIR = "D:\\ayun\\regi\\";
    private static final String LC_SUM_IN = LC_DIR + "lc-sum.xlsx";
    private static final String LC_SUM_OUT = LC_DIR + "lc-sum.txt";
    private static final SimpleDateFormat DATE_SF = new SimpleDateFormat("yyyy-MM-dd");

    public static void main(String[] args) throws IOException {
        run();
    }

    private static final Map<String, Integer> siteSumMap = new HashMap<>();
    private static final Map<String, Integer> dateSumMap = new HashMap<>();

    private static void run() throws IOException {
        extract();
        export();
    }

    private static void export() throws IOException {
        List<String> dateList = new ArrayList<>(dateSumMap.keySet());
        Collections.sort(dateList);

        List<String> siteList = new ArrayList<>(siteSumMap.keySet());
        Collections.sort(siteList);

        StringBuilder txt = new StringBuilder();
        int total = 0;
        for (String date : dateList) {
            int val = dateSumMap.get(date);
            total += val;
            txt.append(date).append("\t").append(val).append("\n");
        }
        txt.append("TOTAL:\t\t").append(total).append("\n");

        txt.append("\n");

        total = 0;
        for (String site : siteList) {
            int val = siteSumMap.get(site);
            total += val;
            txt.append(site).append("\t\t").append(val).append("\n");
        }
        txt.append("TOTAL:\t\t").append(total).append("\n");

        LcUtil.writeFile(LC_SUM_OUT, txt.toString(), 1, true);
    }

    private static void extract() throws FileNotFoundException, IOException {
        FileInputStream fis = new FileInputStream(new File(LC_SUM_IN));
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> ite = sheet.rowIterator();
        boolean finishedAllRow = false;
        boolean head = true;
        while (ite.hasNext()) {
            if (finishedAllRow) {
                break;
            }
            Row row = ite.next();
            Iterator<Cell> cellsInRow = row.cellIterator();

            int colIndex = 0;
            String site = null;
            String dateStr = null;
            while (cellsInRow.hasNext()) {
                if (head) {
                    head = false;
                    break;
                }

                Cell cell = cellsInRow.next();
                System.out.print(cell.toString() + " ");
                if (colIndex == 0) {
                    site = cell.toString().trim();
                    if ("ROW-END".equals(site)) {
                        finishedAllRow = true;
                        break;
                    }
                } else if (colIndex == 1) {
                    try {
                        dateStr = DATE_SF.format(cell.getDateCellValue());
                    } catch (Exception e) {
                        System.err.print(" ERR_DATE ");
                    }
                } else if (colIndex == 2) {
                    String valStr = cell.toString();
                    int val = (int) Double.parseDouble(valStr);
                    if (site != null) {
                        if (!siteSumMap.containsKey(site)) {
                            siteSumMap.put(site, 0);
                        }
                        siteSumMap.put(site, siteSumMap.get(site) + val);
                    }
                    if (dateStr != null) {
                        if (!dateSumMap.containsKey(dateStr)) {
                            dateSumMap.put(dateStr, 0);
                        }
                        dateSumMap.put(dateStr, dateSumMap.get(dateStr) + val);
                    }
                } else {
                    break;
                }
                colIndex++;
            }
            System.out.println();
        }
    }
}
