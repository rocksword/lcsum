package com.an.tool.lcs;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReceiveSummary {
    private static final String LC_DIR = "D:\\ayun\\regi\\";
    private static final String LC_OLD = LC_DIR + "lc-old.xlsx";
    private static final String LC_1610 = LC_DIR + "lc-1610.xlsx";
    private static final String LC_RECEIVE = LC_DIR + "lc-receive.txt";

    private static final SimpleDateFormat DATE_SF = new SimpleDateFormat("yyyy-MM-dd");
    private static Map<String, Double> dateValMap = new HashMap<>();

    public static void main(String[] args) throws IOException {
        run();
    }

    private static void run() throws IOException {
        extractFile(LC_OLD);
        extractFile(LC_1610);

        List<String> dateList = new ArrayList<>(dateValMap.keySet());
        Collections.sort(dateList);

        StringBuilder txt = new StringBuilder();
        int total = 0;
        for (String date : dateList) {
            double dv = dateValMap.get(date);
            total += (int) dv;
            txt.append(date).append("\t").append((int) dv).append("\n");
        }
        txt.append("\n");
        txt.append("TOTAL ").append("\t").append(total).append("\n");
        LcUtil.writeFile(LC_RECEIVE, txt.toString(), 1, true);
    }

    private static Set<String> INVALID_DATE_SET = new HashSet<>();
    static {
        INVALID_DATE_SET.add("DONE");
        INVALID_DATE_SET.add("REGI");
        INVALID_DATE_SET.add("N");
    }

    private static void extractFile(String filePath) throws IOException {
        FileInputStream fis = new FileInputStream(new File(filePath));
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> ite = sheet.rowIterator();

        Set<Integer> dateRowNumSet = new HashSet<>();
        boolean head = true;
        int lastColIndex = -1;
        boolean finishedAllRow = false;
        while (ite.hasNext()) {
            if (finishedAllRow) {
                break;
            }
            Row row = ite.next();
            Iterator<Cell> cellsInRow = row.cellIterator();

            if (head) {
                int colIndex = 0;
                while (cellsInRow.hasNext()) {
                    String cellStr = cellsInRow.next().toString();
                    System.out.print(cellStr + " ");
                    if ("日期".equals(cellStr)) {
                        dateRowNumSet.add(colIndex);
                    } else if ("密码".equals(cellStr)) {
                        lastColIndex = colIndex;
                        System.out.println("Last column index: " + lastColIndex);
                    }
                    colIndex++;
                }
                System.out.println();
                head = false;
                continue;
            }

            int colIndex = 0;
            String val = null;
            while (cellsInRow.hasNext()) {
                if (colIndex == lastColIndex) {
                    break;
                }
                if (colIndex == 16) {
                    System.out.print(" ");
                }

                Cell cell = cellsInRow.next();
                if (colIndex == 1) {
                    if ("点融网".equals(cell.toString())) {
                        System.out.print(" ");
                    }
                    if ("END".equals(cell.toString())) {
                        finishedAllRow = true;
                        break;
                    }
                }

                if (dateRowNumSet.contains(colIndex + 1)) {
                    colIndex++;
                    String cellVal = cell.toString();
                    System.out.print(cellVal + " ");
                    val = cellVal.toString();
                } else if (dateRowNumSet.contains(colIndex)) {
                    colIndex++;
                    if (cell.toString().trim().isEmpty() || INVALID_DATE_SET.contains(cell.toString())) {
                        System.out.print(cell.toString() + " ");
                        continue;
                    }
                    String dateStr = null;
                    try {
                        dateStr = DATE_SF.format(cell.getDateCellValue());
                    } catch (Exception e) {
                        System.out.print(cell.toString() + " ");
                        continue;
                    }

                    if ("1899-12-31".equals(dateStr)) {
                        continue;
                    }
                    System.out.print(dateStr + " ");
                    double tmp = Double.parseDouble(val);
                    if (dateValMap.containsKey(dateStr)) {
                        double v = tmp + dateValMap.get(dateStr);
                        dateValMap.put(dateStr, v);
                    } else {
                        dateValMap.put(dateStr, tmp);
                    }
                } else {
                    colIndex++;
                    String cellVal = cell.toString();
                    System.out.print(cellVal + " ");
                }
            }
            System.out.println();
        }
        fis.close();
    }
}
