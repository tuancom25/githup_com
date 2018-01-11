/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excel.export.ExportFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.AreaPtg;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.RefPtgBase;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.json.JSONObject;

/**
 *
 * @author tuanc
 */
public class Report {

    public static void main(String[] abc) {

        System.out.println(" Hello Excel Report !");
    }
//String result="[{\"SALE_NAME\":\"saler 2\",\"TOTALREVENUE\":163123,\"SALE_USER_NAME\":\"saler2\",\"TOTALPROFIT\":163123,\"TOTAL_CAMP\":5,\"SALE_GROUP_NAME\":\"salegroup 2\",\"TOTAL_ADS\":6,\"TOTALCOST\":0,\"TOTAL_ADVERTISER\":1,\"SALEID\":77,\"DAY\":20171031},{\"SALE_NAME\":\"saler 2\",\"TOTALREVENUE\":223153,\"SALE_USER_NAME\":\"saler2\",\"TOTALPROFIT\":223153,\"TOTAL_CAMP\":5,\"SALE_GROUP_NAME\":"salegroup 2","TOTAL_ADS":6,"TOTALCOST":0,"TOTAL_ADVERTISER":1,"SALEID":77,"DAY":20171030},{"SALE_NAME":"saler 2","TOTALREVENUE":409379,"SALE_USER_NAME":"saler2","TOTALPROFIT":409379,"TOTAL_CAMP":5,"SALE_GROUP_NAME":"salegroup 2","TOTAL_ADS":8,"TOTALCOST":0,"TOTAL_ADVERTISER":1,"SALEID":77,"DAY":20171027},{"SALE_NAME":"saler 2","TOTALREVENUE":0,"SALE_USER_NAME":"saler2","TOTALPROFIT":0,"TOTAL_CAMP":6,"SALE_GROUP_NAME":"salegroup 2","TOTAL_ADS":9,"TOTALCOST":0,"TOTAL_ADVERTISER":1,"SALEID":77,"DAY":20171026},{"SALE_NAME":"saler 2","TOTALREVENUE":140000,"SALE_USER_NAME":"saler2","TOTALPROFIT":140000,"TOTAL_CAMP":2,"SALE_GROUP_NAME":"salegroup 2","TOTAL_ADS":4,"TOTALCOST":0,"TOTAL_ADVERTISER":1,"SALEID":77,"DAY":20171025},{"SALE_NAME":"saler 2","TOTALREVENUE":0,"SALE_USER_NAME":"saler2","TOTALPROFIT":0,"TOTAL_CAMP":2,"SALE_GROUP_NAME":"salegroup 2","TOTAL_ADS":2,"TOTALCOST":0,"TOTAL_ADVERTISER":1,"SALEID":77,"DAY":20171024},{"SALE_NAME":"saler 2","TOTALREVENUE":0,"SALE_USER_NAME":"saler2","TOTALPROFIT":0,"TOTAL_CAMP":5,"SALE_GROUP_NAME":"salegroup 2","TOTAL_ADS":6,"TOTALCOST":0,"TOTAL_ADVERTISER":1,"SALEID":77,"DAY":20171020},{"SALE_NAME":"saler 2","TOTALREVENUE":40000,"SALE_USER_NAME":"saler2","TOTALPROFIT":40000,"TOTAL_CAMP":5,"SALE_GROUP_NAME":"salegroup 2","TOTAL_ADS":5,"TOTALCOST":0,"TOTAL_ADVERTISER":1,"SALEID":77,"DAY":20171018},{"SALE_NAME":"saler 2","TOTALREVENUE":0,"SALE_USER_NAME":"saler2","TOTALPROFIT":0,"TOTAL_CAMP":2,"SALE_GROUP_NAME":"salegroup 2","TOTAL_ADS":2,"TOTALCOST":0,"TOTAL_ADVERTISER":1,"SALEID":77,"DAY":20171017},{"SALE_NAME":"saler 2","TOTALREVENUE":0,"SALE_USER_NAME":"saler2","TOTALPROFIT":0,"TOTAL_CAMP":3,"SALE_GROUP_NAME":"salegroup 2","TOTAL_ADS":3,"TOTALCOST":0,"TOTAL_ADVERTISER":1,"SALEID":77,"DAY":20171016},{"SALE_NAME":null,"TOTALREVENUE":1475655,"SALE_USER_NAME":null,"TOTALPROFIT":1475655,"TOTAL_CAMP":59,"SALE_GROUP_NAME":null,"TOTAL_ADS":70,"TOTALCOST":0,"TOTAL_ADVERTISER":19,"SALEID":null,"DAY":null}]";

//    private Object createReport(String inputTempleteFullFileName, JSONObject oj, JSONArray jsonArray) {
//        Object o = new Object();
//        return o;
//    }
    String file = "";

    void cleanSheet(XSSFSheet sheet) {
        int numberOfRows = sheet.getPhysicalNumberOfRows();
        if (numberOfRows > 0) {
            for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
                if (sheet.getRow(i) != null) {
                    sheet.removeRow(sheet.getRow(i));
                } else {
                    // System.out.println("Info: clean sheet='" + sheet.getSheetName() + "' ... skip line: " + i);
                }
            }
        } else {
            // System.out.println("Info: clean sheet='" + sheet.getSheetName() + "' ... is empty");
        }

    }

    private HashMap<String, String> getMapOfCellss(HSSFSheet sheet, int columnNum, int rowBeginNum) {
        HashMap<String, String> hm = new HashMap<>();
        try {
            //System.out.println("Get Value of Cell =");
            int i = 3;
            String value = "";
            String key = "";
            while (i > 0) {
                try {
                    key = sheet.getRow(i).getCell(1).getStringCellValue();
                    value = sheet.getRow(i).getCell(2).getStringCellValue();
                    hm.put(key, value);
                    // System.out.println(" III " + i);

                    //System.out.println("cell value=" + key + ", value: " + value);
                    if ("null".equals(key) || "".equalsIgnoreCase(key) || i > 150) {
                        break;
                    }
                    i++;
                } catch (Exception x) {
                    System.out.println(" count  fields cellIndex(1) sheet1  : " + i);
                    //  x.printStackTrace();
                    break;
                }
            }

        } catch (Exception x) {
        }
        return hm;
    }

    private HashMap<String, String> getMapOfCellss(XSSFSheet sheet) {
        HashMap<String, String> hm = new HashMap<>();
        try {
            //System.out.println("Get Value of Cell =");
            int i = 3;
            String value = "";
            String key = "";
            while (i > 0) {
                try {
                    key = sheet.getRow(i).getCell(1).getStringCellValue();
                    value = sheet.getRow(i).getCell(2).getStringCellValue();
                    hm.put(key, value);
                    // System.out.println(" III " + i);

                    //System.out.println("cell value=" + key + ", value: " + value);
                    if ("null".equals(key) || "".equalsIgnoreCase(key) || i > 150) {
                        break;
                    }
                    i++;
                } catch (Exception x) {
                    System.out.println(" count  fields cellIndex(1) sheet1  : " + i);
                    //  x.printStackTrace();
                    break;
                }
            }

        } catch (Exception x) {
        }
        return hm;
    }

    private String getHeaderName(XSSFSheet sheet) {
        String name = "";//
        try {
            name = sheet.getRow(1).getCell(1).getStringCellValue();
        } catch (Exception x) {
        }
        return name;
    }

//    private String getTitleName(XSSFSheet sheet) {
//        String name = "";
//        try {
//            name = sheet.getRow(0).getCell(1).getStringCellValue();
//        } catch (Exception x) {
//        }
//        return name;
//    }
//    private String getTitleName(HSSFSheet sheet) {
//        String name = sheet.getRow(0).getCell(1).getStringCellValue();
//        return name;
//    }
    private int getHeaderIndex(XSSFSheet sheet, String headerName) {
        for (int i = 3; i < 100; i++) {
            try {
                String value = sheet.getRow(i).getCell(0).getStringCellValue();
                if (headerName.equalsIgnoreCase(value)) {
                    return i;
                }
            } catch (Exception x) {
            }
        }
        return 0;
    }

//    private int getHeaderIndex(HSSFSheet sheet, String headerName) {
//        for (int i = 3; i < 100; i++) {
//            try {
//                String value = sheet.getRow(i).getCell(0).getStringCellValue();
//                if (headerName.equalsIgnoreCase(value)) {
//                    return i;
//                }
//            } catch (Exception x) {
//            }
//        }
//        return 0;
//    }
    String getFormat() {
        return "";
    }

    protected XSSFWorkbook processExcelFile(String fromDate, String toDate, JSONArray jsonArray, String inputTempleteFullFileNameExcel) {
        String format = getFormat();
        String tongName = "Tổng";
        System.out.println("EXCEL PROCESS  : newProcessExcelFile(HttpServletResponse response()");
        String file = inputTempleteFullFileNameExcel;// myService.getFullFileName(methodname);
        //System.out.println("EXCEL PROCESS :   FullFileName = " + file);
        org.json.JSONArray lsj = jsonArray; // myService.getListObectJson(json, token, methodname);
        int adsRows = lsj.length();
        int adsColumn = 0;
//        String mimeType = "application/octet-stream";
//        String content = "attachment; filename=" + methodname + " _example" + getYYYYMM() + ".xlsx";
//        response.setContentType(mimeType);
//        response.setHeader("Content-disposition", content);
        XSSFCell cellDate = null;

        try {
            File xls = new File(file); // or whatever your file is
            // System.out.println(" Create file complete ");
            FileInputStream in = new FileInputStream(xls);
            //System.out.println(" Create FileInputStream in  ");
            XSSFWorkbook workbook = new XSSFWorkbook(in);
            System.out.println(" Create XSSFWorkbook workbook ");
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFSheet sheet1 = workbook.getSheetAt(1);
            //System.out.println(" Create WookBook and sheet complete ");
            int lastRow = sheet.getLastRowNum();
            //System.out.println(" lastRow= " + lastRow);
            String header = getHeaderName(sheet1);
            int headerIndex = getHeaderIndex(sheet, header);
            XSSFRow headerRow = sheet.getRow(headerIndex);                               // cell(0) sẽ là Số thứ tự. 
            XSSFRow styleRow = sheet.getRow(headerIndex + 1);
            String indexHeader = sheet1.getRow(1).getCell(2).getStringCellValue();
            String indexHeaderDisplay = sheet1.getRow(1).getCell(3).getStringCellValue(); // set this value to cell(0) on headerColumn row

            String fromDateSign = sheet1.getRow(1).getCell(4).getStringCellValue();
            String toDateSign = sheet1.getRow(1).getCell(5).getStringCellValue();
            String fromDateValueInJsonParam = fromDate; //   getAndFormatDate(json, fromDateSign);
            String toDateValueInjsonParam = toDate;// getAndFormatDate(json, toDateSign);
            String titleNameSign = sheet1.getRow(0).getCell(1).getStringCellValue();      //  sheet1.getRow(1).getCell(4).getStringCellValue() ;    getTitleName(sheet1);
            XSSFRow titleRow = sheet.getRow(getHeaderIndex(sheet, titleNameSign));
            for (int i = 0; i < 99; i++) {
                try {
                    cellDate = titleRow.getCell(i);
                    if (i == 0) {
                        cellDate.setCellValue(" ");
                    }
                    String value = cellDate.getStringCellValue();
                    if (value.contains(fromDateSign) || value.contains(toDateSign)) {
                        String t = value.replace(fromDateSign, fromDateValueInJsonParam).replace(toDateSign, toDateValueInjsonParam);
                        cellDate.setCellValue(t);
                        break;
                    }
                } catch (Exception x) {
                }
            }

            HashMap<String, String> hm = getMapOfCellss(sheet1);
            adsColumn = hm.size();
            //System.out.println("HM  complete !");
            sheet.shiftRows(headerIndex + 1, lastRow + 1, adsRows + 1);
            //System.out.println("ShiftRow complete !");                   
            for (int j = 0; j < adsRows; j++) {
                try {
                    XSSFRow row1 = sheet.createRow(headerIndex + 1 + j);
                    org.json.JSONObject ojs = lsj.getJSONObject(j);
                    //System.out.println("Test row value: " + ojs.toString());
                    for (int jj = 0; jj <= adsColumn + 5; jj++) {
                        try {
                            //System.out.println("j=" + j + ", jj= " + jj);
                            if (j == adsRows - 1 & jj == 1) {

                                XSSFCell cell2 = row1.createCell(jj);
                                // set "Tổng :" into cell 
                                cell2.setCellValue(tongName);
                                //System.out.println("" + ojs.toString());
                            } else {
                                //System.out.println("If ELSE :  j=" + j + ", jj= " + jj);
//                                if(jj==0){ 
//                                            cell2.setCellValue(j+1); 
//                                }
                                Object o = null;
                                try {
                                    //System.out.println("If ELSE :  j=" + j + ", jj= " + jj);
                                    //System.out.println("Test headerRow Cell: " + jj + " , getcell(1) " + headerRow.getCell(1).getStringCellValue()+ ", getcell(2) ");
                                    String m = "" + headerRow.getCell(jj).getStringCellValue();
                                    // System.out.println("XXXXYYYMMMM  jsonkey mm: " + indexHeader + ", m = " + m);  
                                    // System.out.println("YYYYYYYY  indexheader : " + indexHeader + " - jsonkey: " + m  + ", " +  ",  column: " + jj);     
                                    String jsonkey = m;
                                    XSSFCellStyle cellStyle = null;
                                    try {
                                        cellStyle = styleRow.getCell(jj).getCellStyle();
                                    } catch (Exception x) {
                                    }

                                    if (indexHeader.toLowerCase().equalsIgnoreCase(m) && j < (adsRows - 1)) {
                                        // System.out.println("   XXXXXXXXX   o = " + j);
                                        // o = j;                                // set so thu tu theo y' muon. 
                                    } else {
                                        o = ojs.get(jsonkey);
                                    }
                                    if (jsonkey.toUpperCase().contains("DAY")) {
                                        // o = convertDate(format, o.toString());
                                    }
                                    XSSFCell cell2 = row1.createCell(jj);
                                    if (cellStyle != null) {
                                        cell2.setCellStyle(cellStyle);
                                    }
                                    if (o instanceof Integer) {
                                        // check set  so thu tu, index  excel. 
                                        Integer b = (Integer) o;
                                        cell2.setCellValue(b);
                                    } else if (o instanceof Number) {
                                        Double b = (Double) o;
                                        cell2.setCellValue(b);
                                    } else if (o instanceof String) {
                                        String tem = (String) o;
                                        cell2.setCellValue(tem);
                                    } else if (o instanceof java.util.Date) {
                                        cell2.setCellValue((Date) o);
                                    } else if (o == null) {

                                        cell2.setCellValue("");
                                    } else {
                                        String temp = "" + o.toString().toLowerCase();
                                        if (temp.contains("null")) {
                                            cell2.setCellValue(" ");
                                        }
                                    }
                                } catch (Exception c) {

                                    //  System.out.println(" Convert data type is false  value=" + o + " , index column: " + jj);
                                }
                            }
                        } catch (Exception x) {
                        }
                    }
                } catch (Exception v) {
                }
            }
            // reset header display 
            for (int n = 0; n <= adsColumn + 5; n++) {
                try {
                    if (n == 0) {
                        headerRow.getCell(0).setCellValue(" ");
                    } else {
                        String m = headerRow.getCell(n).getStringCellValue();
                        String headerTemp = "";
                        String nn = "";
                        if (m != null && !("".equals(m))) {
//                           Hashtable<String,String> htm = new Hashtable() ;
//                            Enumeration<String> keyss=   htm.keys();
                            Set<String> keyset = hm.keySet();
                            for (String tempKey : keyset) {
                                if (m.contains(tempKey)) {
                                    nn = hm.get(tempKey);
                                    headerTemp = m.replace(tempKey, nn);
                                }
                            }
                            // System.out.println("headerIndex:"+ m + ", header display: " + nn+ ", headerIndex Display: " + indexHeaderDisplay);
                            if (m.contains(indexHeader)) {
                                //  headerRow.getCell(n).setCellValue(indexHeaderDisplay);   // da set index la column(0). 
                            } else {
                                headerRow.getCell(n).setCellValue(headerTemp);
                            }
                        }
                    }
                } catch (Exception c) {

                }
            }
            // for loop for set index-(so thu tu trong bang tinh) .
            for (int i = 0; i < adsRows; i++) {
                try {
                    // System.out.println("set value " +i);
                    XSSFRow r = sheet.getRow(i + headerIndex);
                    XSSFCell c = r.createCell(0);
                    if (i == 0) {
                        c.setCellValue(indexHeaderDisplay);
                        continue;
                    }
                    c.setCellValue(i);
                } catch (Exception x) {
                    System.out.println("set value fail  " + i + ", " + x.getMessage());
                }
            }

            cleanSheet(sheet1);
            sheet.removeRow(styleRow);
            // clearColumn(sheet, 0);
            //FileOutputStream fos =new FileOutputStream(new File("H:/test/test.xlsx"));
            //workbook.write(fos);

            System.out.println("BEGIN write file Response");
            System.out.println("FILE  Done Response ");
            return workbook;
        } catch (IOException x) {
            System.out.println("LoI Thao tac file ");
            x.printStackTrace();

        }
        return null;
    }

    /**
     * /* tuancom25@gmail.com /
     *
     * @param fromDate date select begin at report
     * @param toDate date select end in report
     * @param curentDate
     * @param email
     * @param accountName
     * @param jsonArray : tương đương bảng data.
     * @param inputTempleteFileExcel you need close your inputTemleteFileExcel =
     * new InputStream(fileName);
     * @param tong / Danh sách các tham số này là cụ thể /
     * @return workbook ;
     *
     *
     */
    public XSSFWorkbook processExcelFile(String fromDate, String toDate, String curentDate, String email, String accountName, JSONArray jsonArray, FileInputStream inputTempleteFileExcel, String tong) {
        String format = getFormat();
        String tongName = tong;// "Tổng" ;
        System.out.println("EXCEL PROCESS  : newProcessExcelFile(HttpServletResponse response()");
        org.json.JSONArray lsj = jsonArray; // myService.getListObectJson(json, token, methodname);
        int adsRows = lsj.length();
        int adsColumn = 0;
        try {

            XSSFWorkbook workbook = new XSSFWorkbook(inputTempleteFileExcel);
            System.out.println("XXXX Create XSSFWorkbook workbook  ");
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFSheet sheet1 = workbook.getSheetAt(1);
            //System.out.println(" Create WookBook and sheet complete ");
            int lastRow = sheet.getLastRowNum();
            //System.out.println(" lastRow= " + lastRow);
            String header = getHeaderName(sheet1);
            int headerIndex = getHeaderIndex(sheet, header);
            for (int rowIndex = 0; rowIndex < headerIndex - 1; rowIndex++) {
                XSSFRow r = sheet.getRow(rowIndex);
                for (int cindex = 0; cindex < 32; cindex++) {
                    try {
                        XSSFCell cl = r.getCell(cindex);
                        String value = cl.getStringCellValue();
                        //System.out.println("Value = " + value + ", cindex = " + cindex + ", rowindex= " + rowIndex);
                        String temp = "";
                        if (value.toLowerCase().contains("$email$")) {
                            temp = value.replace("$email$", email);
                            cl.setCellValue(temp);
                        }
                        if (value.toLowerCase().contains("$accountname$")) {
                            temp = value.replace("$accountname$", accountName).replace("$accountName$", accountName);
                            cl.setCellValue(temp);
                        }
                        if (value.toLowerCase().contains("$currentdate$")) {
                            temp = value.replace("$currentdate$", curentDate).replace("$currentDate$", curentDate);
                            cl.setCellValue(temp);
                        }
                        if (value.toLowerCase().contains("$fromdate$")) {
                            temp = value.replace("$fromdate$", fromDate).replace("$fromDate$", fromDate).replace("$todate$", toDate).replace("$toDate$", toDate);
                            cl.setCellValue(temp);
                        }

                    } catch (Exception x) {
                    }

                }
            }
            XSSFRow headerRow = sheet.getRow(headerIndex);                               // cell(0) sẽ là Số thứ tự.
            XSSFRow formulaRow = sheet.getRow(headerIndex - 1);
            XSSFRow styleRow = sheet.getRow(headerIndex + 1);
            formulaRow = styleRow;                                      // dịch chuyển  con trỏ row xuống hàng sytle. hàng này vừa đóng vai trò là style vừa đóng vai trò là công thức.
            String indexHeader = sheet1.getRow(1).getCell(2).getStringCellValue();
            String indexHeaderDisplay = sheet1.getRow(1).getCell(3).getStringCellValue(); // set this value to cell(0) on headerColumn row
            String formulaString = "";
            XSSFCell cellsrc = null;

            HashMap<String, String> hm = getMapOfCellss(sheet1);
            adsColumn = hm.size();
            System.out.println("HashMap  complete !");
            sheet.shiftRows(headerIndex + 1, lastRow + 1, adsRows + 1);
            System.out.println("ShiftRow complete ! BEGIN fill data  ");
            for (int j = 0; j < adsRows; j++) {
                try {
                    XSSFRow row1 = sheet.createRow(headerIndex + 1 + j);
                    org.json.JSONObject ojs = lsj.getJSONObject(j);
                    //System.out.println("Test row value: " + ojs.toString());
                    for (int jj = 0; jj <= adsColumn + 5; jj++) {                // hoặc  jj < 26 nhở hơn số các ký tụ của cột, a,b,c,d, ... z 26 ký tự 
                        try {
                            //System.out.println("j=" + j + ", jj= " + jj);
                            try {
                                cellsrc = formulaRow.getCell(jj);
                                // System.out.println(" Value :onformula   "  + cellsrc.getRawValue());                               
                                formulaString = cellsrc.getCellFormula();
                                if (formulaString != null) {
//                                    System.out.println("formula: " + formulaString);
                                } else {
                                }
                            } catch (Exception c) {
                            }
                            if (j == adsRows - 1 & jj == 1) {

                                XSSFCell cell2 = row1.createCell(jj);
                                // set "Tổng :" into cell 
                                cell2.setCellValue(tongName);
                                //System.out.println("" + ojs.toString());
                            } else {

                                Object o = null;
                                try {
                                    //System.out.println("Test headerRow Cell: " + jj + " , getcell(1) " + headerRow.getCell(1).getStringCellValue()+ ", getcell(2) ");
                                    String m = "" + headerRow.getCell(jj).getStringCellValue();
                                    String n = "";

                                    // System.out.println("YYYYYYYY  indexheader : " + indexHeader + " - jsonkey: " + m  + ", " +  ",  column: " + jj);     
                                    String jsonkey = m;
                                    XSSFCellStyle cellStyle = null;
                                    try {
                                        cellStyle = styleRow.getCell(jj).getCellStyle();
                                    } catch (Exception x) {
                                    }

                                    if (indexHeader.toLowerCase().equalsIgnoreCase(m) && j < (adsRows - 1)) {
                                        // System.out.println("   XXXXXXXXX   o = " + j);
                                        // o = j;                                // set so thu tu theo y' muon. 
                                    } else {
                                        try {
                                            o = ojs.get(jsonkey);
                                        } catch (Exception x) {
                                        }
                                    }
                                    if (jsonkey.toUpperCase().contains("DAY")) {
                                        try {
                                            // o = convertDate(format, o.toString());
                                        } catch (Exception x) {
                                        }
                                    }
                                    XSSFCell cell2 = row1.createCell(jj);
                                    if (cellStyle != null && j < adsRows - 1) {
                                        cell2.setCellStyle(cellStyle);
                                    }
                                    if (o instanceof Integer) {
                                        // check set  so thu tu, index  excel. 
                                        Integer b = (Integer) o;
                                        cell2.setCellValue(b);
                                    } else if (o instanceof Number) {
                                        Double b = (Double) o;
                                        cell2.setCellValue(b);
                                    } else if (o instanceof String) {
                                        String tem = (String) o;
                                        cell2.setCellValue(tem);
                                    } else if (o instanceof java.util.Date) {
                                        cell2.setCellValue((Date) o);
                                    } else if (o == null) {

                                        if (formulaString != null) {
                                        } else {
                                            cell2.setCellValue("");
                                        }
                                    } else {
                                        String temp = "" + o.toString().toLowerCase();
                                        if (temp.contains("null")) {
                                            cell2.setCellValue(" ");
                                        }
                                    }

                                    if (formulaString != null) {
                                        // System.out.println(" Begin Copy formula xxxxxxx formulaString = " + formulaString);
                                        int shiftRows = cell2.getRowIndex() - cellsrc.getRowIndex();
                                        int shiftCols = cell2.getColumnIndex() - cellsrc.getColumnIndex();
                                        //copyFormula(workbook, formulaString, cellsrc, cell2);
                                        String tempFormula = getFormula(workbook, formulaString, shiftRows, shiftCols);
                                        //System.out.println("String tempFormula = getFormula(workbook, formulaString, shiftRows, shiftCols) tempFormula= " + tempFormula);
//                                        if (cell2 == null) {
//                                            System.out.println("cell2 is null ");
//                                        } else {
//                                            System.out.println("cell 2 not null ");
//                                        }
                                        cell2.setCellFormula(tempFormula);
                                        //                                       System.out.println("-== Set formul ok. String tempFormula = getFormula(workbook, formulaString, shiftRows, shiftCols) tempFormula= " + tempFormula);
                                        formulaString = null;
                                    } else {
                                    }

                                } catch (Exception c) {
                                    //  System.out.println(" Convert data type is false  value=" + o + " , index column: " + jj);
                                }
                            }
                        } catch (Exception x) {
                        }
                    }
                } catch (Exception v) {
                }
            }

            // reset header display
            Set<String> keyset = hm.keySet();
            for (int n = 0; n <= adsColumn + 5; n++) {
                try {
                    if (n == 0) {
                        headerRow.getCell(0).setCellValue(" ");
                    } else {
                        String m = headerRow.getCell(n).getStringCellValue();
                        String headerTemp = "";
                        String nn = "";
                        // boolean contentHeader =false;
                        if (m != null && !("".equals(m))) {
//                           Hashtable<String,String> htm = new Hashtable() ;                           
                            for (String tempKey : keyset) {
                                if (m.contains(tempKey)) {
                                    nn = hm.get(tempKey);
                                    headerTemp = m.replace(tempKey, nn);
                                    // contentHeader=true;
                                    break;
                                }

                            }

                            if (headerTemp != null && headerTemp.length() > 1) {
                                headerRow.getCell(n).setCellValue(headerTemp);   // Tạm cất không thay đổi nội dung tiêu đề.  
                            }
                        }
                    }
                } catch (Exception c) {

                }
            }
            // for loop for set index-(so thu tu trong bang tinh) .
            for (int i = 0; i < adsRows; i++) {
                try {
                    // System.out.println("set value " +i);
                    XSSFRow r = sheet.getRow(i + headerIndex);
                    XSSFCell c = r.createCell(0);
                    if (i == 0) {
                        c.setCellValue(indexHeaderDisplay);
                        continue;
                    }
                    c.setCellValue(i);
                } catch (Exception x) {
                    System.out.println("set value fail  " + i + ", " + x.getMessage());
                }
            }

            cleanSheet(sheet1);
            sheet.removeRow(styleRow);
            //sheet.removeRow(formulaRow);
            //FileOutputStream fos =new FileOutputStream(new File("H:/test/test.xlsx"));
            //workbook.write(fos);

            System.out.println("BEGIN write file Response");
            //  workbook.write(response.getOutputStream());
            //  workbook.close();
            //  in.close();
            System.out.println("FILE  Done Response ");
            return workbook;
        } catch (IOException x) {
            System.out.println("LoI Thao tac file ");
            x.printStackTrace();

        }
        return null;
    }

    /**
     *
     * @param workbook
     * @param formula
     * @param shiftRows
     * @param shiftCols
     * @return
     */
    private String getFormula(XSSFWorkbook workbook, String formula, int shiftRows, int shiftCols) {
        //System.out.println(" XXXX   cscdcdscd                  ÂSASAS");
        String tempFormula = "";
        try {
            XSSFEvaluationWorkbook workbookWrapper = XSSFEvaluationWorkbook.create(workbook);
            Ptg[] ptgs = FormulaParser.parse(formula, workbookWrapper, FormulaType.CELL, 0);
            //System.out.println(" XXXX   YYYYYYYYY    HHHHHHHHHHHHHHHHH   formula= " + formula);
            for (Ptg ptg : ptgs) {
                if (ptg instanceof RefPtgBase) // base class for cell references
                {
                    RefPtgBase ref = (RefPtgBase) ptg;
                    if (ref.isColRelative()) {
                        ref.setColumn(ref.getColumn() + shiftCols);
                    }
                    if (ref.isRowRelative()) {
                        ref.setRow(ref.getRow() + shiftRows);
                    }
                } else if (ptg instanceof AreaPtg) // base class for range references
                {
                    AreaPtg ref = (AreaPtg) ptg;
                    if (ref.isFirstColRelative()) {
                        ref.setFirstColumn(ref.getFirstColumn() + shiftCols);
                    }
                    if (ref.isLastColRelative()) {
                        ref.setLastColumn(ref.getLastColumn() + shiftCols);
                    }
                    if (ref.isFirstRowRelative()) {
                        ref.setFirstRow(ref.getFirstRow() + shiftRows);
                    }
                    if (ref.isLastRowRelative()) {
                        ref.setLastRow(ref.getLastRow() + shiftRows);
                    }
                }
            }
            tempFormula = FormulaRenderer.toFormulaString(workbookWrapper, ptgs);
            //System.out.println("() 0000 111 tempFormula= " + tempFormula);
        } catch (Exception x) {
        }
        return tempFormula;
    }

    /**
     * @param paramaters : HashMap<String,String> Danh sách các tham số phía
     * tiêu đề của excel.
     * @param jsonArray : List Object json data
     * @param inputTempleteFileExcel: là 1 excel dạng FileInputStream
     * @param tong : Tham số hiển thị phía phía cuối, hàng tổng của báo cáo.
     * @return Trả về 1 workbook để viết lên luồng xuất file, file ổ cứng, hoặc
     * luồng gửi client responses.getOutputStream().
     *
     */
    public XSSFWorkbook processExcelFile(HashMap<String, String> paramaters, JSONArray jsonArray, FileInputStream inputTempleteFileExcel, String tong) {
        // HashMap<String,String> paramaters =null;
        //String format = getFormat();
        String tongName = tong;// "Tổng" ;
        System.out.println("EXCEL PROCESS  : newProcessExcelFile(HttpServletResponse response()");
        org.json.JSONArray lsj = jsonArray; // myService.getListObectJson(json, token, methodname);
        int adsRows = lsj.length();
        int adsColumn = 0;
        Set<String> keyss = paramaters.keySet();
        try {

            XSSFWorkbook workbook = new XSSFWorkbook(inputTempleteFileExcel);
            System.out.println("XXXX Create XSSFWorkbook workbook  ");
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFSheet sheet1 = workbook.getSheetAt(1);
            //System.out.println(" Create WookBook and sheet complete ");
            int lastRow = sheet.getLastRowNum();
            //System.out.println(" lastRow= " + lastRow);
            String header = getHeaderName(sheet1);
            Boolean setValue = false;
            int headerIndex = getHeaderIndex(sheet, header);
            for (int rowIndex = 0; rowIndex < headerIndex - 1; rowIndex++) {
                XSSFRow r = sheet.getRow(rowIndex);
                for (int cindex = 0; cindex < 32; cindex++) {
                    try {
                        XSSFCell cl = r.getCell(cindex);
                        String value = cl.getStringCellValue();
                        if (value != null) {
                            String temp = "";
                            for (String s : keyss) {
                                try {
                                    if (value.toLowerCase().contains(s.toLowerCase())) {

                                        String ss = "" + paramaters.get(s);
                                        temp = value.replace(s, ss).replace("$", "");
                                        System.out.println("EXCEL PROCESS: CellValue = " + value + ", cellindex = " + cindex + ", rowindex= " + rowIndex + ", Hashmap key = " + s + ", HashMap value=" + ss);
                                        value = temp;
                                        //System.out.println("data temp = " + temp+  ", ss = " + ss);
                                        setValue = true;
                                    }
                                } catch (Exception x) {
                                }
                            }
                            if (setValue) {
                                System.out.println("Data temp BEBORE UPDATE = " + temp);
                                cl.setCellValue(temp.toString());
                            }
                            setValue = false;
                        }

                    } catch (Exception x) {
                    }

                }
            }
            XSSFRow headerRow = sheet.getRow(headerIndex);                               // cell(0) sẽ là Số thứ tự.
            // XSSFRow formulaRow = sheet.getRow(headerIndex - 1);
            XSSFRow styleRow = sheet.getRow(headerIndex + 1);
            XSSFRow formulaRow = styleRow;                                      // dịch chuyển  con trỏ row xuống hàng sytle. hàng này vừa đóng vai trò là style vừa đóng vai trò là công thức.
            String indexHeader = sheet1.getRow(1).getCell(2).getStringCellValue();
            String indexHeaderDisplay = sheet1.getRow(1).getCell(3).getStringCellValue(); // set this value to cell(0) on headerColumn row
            String formulaString = "";
            XSSFCell cellsrc = null;

            HashMap<String, String> hm = getMapOfCellss(sheet1);
            adsColumn = hm.size();
            System.out.println("HashMap  complete !");
            sheet.shiftRows(headerIndex + 1, lastRow + 1, adsRows + 1);
            System.out.println("ShiftRow complete ! BEGIN fill data  ");
            for (int j = 0; j < adsRows; j++) {
                try {
                    XSSFRow row1 = sheet.createRow(headerIndex + 1 + j);
                    org.json.JSONObject ojs = lsj.getJSONObject(j);
                    //System.out.println("Test row value: " + ojs.toString());
                    for (int jj = 0; jj <= adsColumn + 5; jj++) {                // hoặc  jj < 26 nhở hơn số các ký tụ của cột, a,b,c,d, ... z 26 ký tự 
                        try {
                            //System.out.println("j=" + j + ", jj= " + jj);
                            try {
                                cellsrc = formulaRow.getCell(jj);
                                // System.out.println(" Value :onformula   "  + cellsrc.getRawValue());                               
                                formulaString = cellsrc.getCellFormula();
                                if (formulaString != null) {
//                                    System.out.println("formula: " + formulaString);
                                } else {
                                }
                            } catch (Exception c) {
                            }
                            if (j == adsRows - 1 & jj == 1) {

                                XSSFCell cell2 = row1.createCell(jj);
                                // set "Tổng :" into cell 
                                cell2.setCellValue(tongName);
                                //System.out.println("" + ojs.toString());
                            } else {

                                Object o = null;
                                try {
                                    //System.out.println("Test headerRow Cell: " + jj + " , getcell(1) " + headerRow.getCell(1).getStringCellValue()+ ", getcell(2) ");
                                    String m = "" + headerRow.getCell(jj).getStringCellValue();
                                    String n = "";

                                    // System.out.println("YYYYYYYY  indexheader : " + indexHeader + " - jsonkey: " + m  + ", " +  ",  column: " + jj);     
                                    String jsonkey = m;
                                    XSSFCellStyle cellStyle = null;
                                    try {
                                        cellStyle = styleRow.getCell(jj).getCellStyle();
                                    } catch (Exception x) {
                                    }

                                    if (indexHeader.toLowerCase().equalsIgnoreCase(m) && j < (adsRows - 1)) {
                                        // System.out.println("   XXXXXXXXX   o = " + j);
                                        // o = j;                                // set so thu tu theo y' muon. 
                                    } else {
                                        try {
                                            o = ojs.get(jsonkey);
                                        } catch (Exception x) {
                                        }
                                    }
                                    if (jsonkey.toUpperCase().contains("DAY")) {
                                        try {
                                            // o = convertDate(format, o.toString());
                                        } catch (Exception x) {
                                        }
                                    }
                                    XSSFCell cell2 = row1.createCell(jj);
                                    if (cellStyle != null && j < adsRows - 1) {
                                        cell2.setCellStyle(cellStyle);
                                    }
                                    if (o instanceof Integer) {
                                        // check set  so thu tu, index  excel. 
                                        Integer b = (Integer) o;
                                        cell2.setCellValue(b);
                                    } else if (o instanceof Number) {
                                        Double b = (Double) o;
                                        cell2.setCellValue(b);
                                    } else if (o instanceof String) {
                                        String tem = (String) o;
                                        cell2.setCellValue(tem);
                                    } else if (o instanceof java.util.Date) {
                                        cell2.setCellValue((Date) o);
                                    } else if (o == null) {

                                        if (formulaString != null) {
                                        } else {
                                            cell2.setCellValue("");
                                        }
                                    } else {
                                        String temp = "" + o.toString().toLowerCase();
                                        if (temp.contains("null")) {
                                            cell2.setCellValue(" ");
                                        }
                                    }

                                    if (formulaString != null) {
                                        // System.out.println(" Begin Copy formula xxxxxxx formulaString = " + formulaString);
                                        int shiftRows = cell2.getRowIndex() - cellsrc.getRowIndex();
                                        int shiftCols = cell2.getColumnIndex() - cellsrc.getColumnIndex();
                                        //copyFormula(workbook, formulaString, cellsrc, cell2);
                                        String tempFormula = getFormula(workbook, formulaString, shiftRows, shiftCols);
                                        //System.out.println("String tempFormula = getFormula(workbook, formulaString, shiftRows, shiftCols) tempFormula= " + tempFormula);
//                                        if (cell2 == null) {
//                                            System.out.println("cell2 is null ");
//                                        } else {
//                                            System.out.println("cell 2 not null ");
//                                        }
                                        cell2.setCellFormula(tempFormula);
                                        //                                       System.out.println("-== Set formul ok. String tempFormula = getFormula(workbook, formulaString, shiftRows, shiftCols) tempFormula= " + tempFormula);
                                        formulaString = null;
                                    } else {
                                    }

                                } catch (Exception c) {
                                    //  System.out.println(" Convert data type is false  value=" + o + " , index column: " + jj);
                                }
                            }
                        } catch (Exception x) {
                        }
                    }
                } catch (Exception v) {
                }
            }

            // reset header display
            Set<String> keyset = hm.keySet();
            for (int n = 0; n <= adsColumn + 5; n++) {
                try {
                    if (n == 0) {
                        headerRow.getCell(0).setCellValue(" ");
                    } else {
                        String m = headerRow.getCell(n).getStringCellValue();
                        String headerTemp = "";
                        String nn = "";
                        // boolean contentHeader =false;
                        if (m != null && !("".equals(m))) {
//                           Hashtable<String,String> htm = new Hashtable() ;                           
                            for (String tempKey : keyset) {
                                if (m.contains(tempKey)) {
                                    nn = hm.get(tempKey);
                                    headerTemp = m.replace(tempKey, nn);
                                    // contentHeader=true;
                                    break;
                                }

                            }

                            if (headerTemp != null && headerTemp.length() > 1) {
                                headerRow.getCell(n).setCellValue(headerTemp);   // Tạm cất không thay đổi nội dung tiêu đề.  
                            }
                        }
                    }
                } catch (Exception c) {

                }
            }
            // for loop for set index-(so thu tu trong bang tinh) .
            for (int i = 0; i < adsRows; i++) {
                try {
                    // System.out.println("set value " +i);
                    XSSFRow r = sheet.getRow(i + headerIndex);
                    XSSFCell c = r.createCell(0);
                    if (i == 0) {
                        c.setCellValue(indexHeaderDisplay);
                        continue;
                    }
                    c.setCellValue(i);
                } catch (Exception x) {
                    System.out.println("set value fail  " + i + ", " + x.getMessage());
                }
            }

            cleanSheet(sheet1);
            sheet.removeRow(styleRow);
            //sheet.removeRow(formulaRow);
            //FileOutputStream fos =new FileOutputStream(new File("H:/test/test.xlsx"));
            //workbook.write(fos);

            System.out.println("BEGIN write file Response");
            //  workbook.write(response.getOutputStream());
            //  workbook.close();
            //  in.close();
            System.out.println("FILE  Done Response ");
            return workbook;
        } catch (IOException x) {
            System.out.println("LoI Thao tac file ");
            x.printStackTrace();

        }
        return null;
    }

    public XSSFWorkbook processExcelFile(JSONObject paramaters, JSONArray jsonArray, FileInputStream inputTempleteFileExcel, String tong) {
        // HashMap<String,String> paramaters =null;
        //String format = getFormat();
        String tongName = tong;// "Tổng" ;
        System.out.println("EXCEL PROCESS  : newProcessExcelFile(HttpServletResponse response()");
        org.json.JSONArray lsj = jsonArray; // myService.getListObectJson(json, token, methodname);
        int adsRows = lsj.length();
        int adsColumn = 0;
        Set<String> keyss = paramaters.keySet();
        System.out.println(" ");
        try {

            XSSFWorkbook workbook = new XSSFWorkbook(inputTempleteFileExcel);
            System.out.println("XXXX Create XSSFWorkbook workbook  ");
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFSheet sheet1 = workbook.getSheetAt(1);
            //System.out.println(" Create WookBook and sheet complete ");
            int lastRow = sheet.getLastRowNum();
            //System.out.println(" lastRow= " + lastRow);
            String header = getHeaderName(sheet1);
            int headerIndex = getHeaderIndex(sheet, header);
            boolean setValue = false;
            for (int rowIndex = 0; rowIndex < headerIndex - 1; rowIndex++) {
                XSSFRow r = sheet.getRow(rowIndex);
                for (int cindex = 0; cindex < 32; cindex++) {
                    try {
                        XSSFCell cl = r.getCell(cindex);
                        String value = cl.getStringCellValue();
                        if (value != null && value.trim().length() > 1) {
                            String temp = "";

                            for (String s : keyss) {
                                try {
                                    if (value.toLowerCase().contains(s.toLowerCase())) {
                                        String ss = "" + paramaters.get(s);
                                        temp = value.replace(s, ss).replace("$", "");
                                        System.out.println("EXCELL PROCESS: CellValue = " + value + ", cellindex = " + cindex + ", rowindex= " + rowIndex + ". Data temp=" + temp + ", json key = " + s + ", json value=" + ss);
                                        value = temp;
                                        //  System.out.println("data temp = " + temp+  ", ss = " + ss);
                                        setValue = true;
                                    }
                                } catch (Exception x) {
                                }
                            }
                            if (setValue) {
                                System.out.println("Data temp BEBORE UPDATE CELL = " + temp);
                                cl.setCellValue(temp.toString());
                            }
                            setValue = false;
                        }

                    } catch (Exception x) {
                    }

                }
            }
            XSSFRow headerRow = sheet.getRow(headerIndex);                               // cell(0) sẽ là Số thứ tự.
            // XSSFRow formulaRow = sheet.getRow(headerIndex - 1);
            XSSFRow styleRow = sheet.getRow(headerIndex + 1);
            XSSFRow formulaRow = styleRow;                                      // dịch chuyển  con trỏ row xuống hàng sytle. hàng này vừa đóng vai trò là style vừa đóng vai trò là công thức.
            String indexHeader = sheet1.getRow(1).getCell(2).getStringCellValue();
            String indexHeaderDisplay = sheet1.getRow(1).getCell(3).getStringCellValue(); // set this value to cell(0) on headerColumn row
            String formulaString = "";
            XSSFCell cellsrc = null;

            HashMap<String, String> hm = getMapOfCellss(sheet1);
            adsColumn = hm.size();
            System.out.println("HashMap  complete !");
            sheet.shiftRows(headerIndex + 1, lastRow + 1, adsRows + 1);
            System.out.println("ShiftRow complete ! BEGIN fill data  ");
            for (int j = 0; j < adsRows; j++) {
                try {
                    XSSFRow row1 = sheet.createRow(headerIndex + 1 + j);
                    org.json.JSONObject ojs = lsj.getJSONObject(j);
                    //System.out.println("Test row value: " + ojs.toString());
                    for (int jj = 0; jj <= adsColumn + 5; jj++) {                // hoặc  jj < 26 nhở hơn số các ký tụ của cột, a,b,c,d, ... z 26 ký tự 
                        try {
                            //System.out.println("j=" + j + ", jj= " + jj);
                            try {
                                cellsrc = formulaRow.getCell(jj);
                                // System.out.println(" Value :onformula   "  + cellsrc.getRawValue());                               
                                formulaString = cellsrc.getCellFormula();
                                if (formulaString != null) {
                                    //System.out.println("formula: " + formulaString);
                                } else {
                                }
                            } catch (Exception c) {
                            }
                            if (j == adsRows - 1 & jj == 1) {

                                XSSFCell cell2 = row1.createCell(jj);
                                // set "Tổng :" into cell 
                                cell2.setCellValue(tongName);
                                //System.out.println("" + ojs.toString());
                            } else {

                                Object o = null;
                                try {
                                    //System.out.println("Test headerRow Cell: " + jj + " , getcell(1) " + headerRow.getCell(1).getStringCellValue()+ ", getcell(2) ");
                                    String m = "" + headerRow.getCell(jj).getStringCellValue();
                                    String n = "";

                                    // System.out.println("YYYYYYYY  indexheader : " + indexHeader + " - jsonkey: " + m  + ", " +  ",  column: " + jj);     
                                    String jsonkey = m;
                                    XSSFCellStyle cellStyle = null;
                                    try {
                                        cellStyle = styleRow.getCell(jj).getCellStyle();
                                    } catch (Exception x) {
                                    }

                                    if (indexHeader.toLowerCase().equalsIgnoreCase(m) && j < (adsRows - 1)) {
                                        // System.out.println("   XXXXXXXXX   o = " + j);
                                        // o = j;                                // set so thu tu theo y' muon. 
                                    } else {
                                        try {
                                            o = ojs.get(jsonkey);
                                        } catch (Exception x) {
                                        }
                                    }
                                    if (jsonkey.toUpperCase().contains("DAY")) {
                                        try {
                                            // o = convertDate(format, o.toString());
                                        } catch (Exception x) {
                                        }
                                    }
                                    XSSFCell cell2 = row1.createCell(jj);
                                    if (cellStyle != null && j < adsRows - 1) {
                                        cell2.setCellStyle(cellStyle);
                                    }
                                    if (o instanceof Integer) {
                                        // check set  so thu tu, index  excel. 
                                        Integer b = (Integer) o;
                                        cell2.setCellValue(b);
                                    } else if (o instanceof Number) {
                                        Double b = (Double) o;
                                        cell2.setCellValue(b);
                                    } else if (o instanceof String) {
                                        String tem = (String) o;
                                        cell2.setCellValue(tem);
                                    } else if (o instanceof java.util.Date) {
                                        cell2.setCellValue((Date) o);
                                    } else if (o == null) {

                                        if (formulaString != null) {
                                        } else {
                                            cell2.setCellValue("");
                                        }
                                    } else {
                                        String temp = "" + o.toString().toLowerCase();
                                        if (temp.contains("null")) {
                                            cell2.setCellValue(" ");
                                        }
                                    }

                                    if (formulaString != null) {
                                        // System.out.println(" Begin Copy formula xxxxxxx formulaString = " + formulaString);
                                        int shiftRows = cell2.getRowIndex() - cellsrc.getRowIndex();
                                        int shiftCols = cell2.getColumnIndex() - cellsrc.getColumnIndex();
                                        //copyFormula(workbook, formulaString, cellsrc, cell2);
                                        String tempFormula = getFormula(workbook, formulaString, shiftRows, shiftCols);
                                        //System.out.println("String tempFormula = getFormula(workbook, formulaString, shiftRows, shiftCols) tempFormula= " + tempFormula);
//                                        if (cell2 == null) {
//                                            System.out.println("cell2 is null ");
//                                        } else {
//                                            System.out.println("cell 2 not null ");
//                                        }
                                        cell2.setCellFormula(tempFormula);
                                        //                                       System.out.println("-== Set formul ok. String tempFormula = getFormula(workbook, formulaString, shiftRows, shiftCols) tempFormula= " + tempFormula);
                                        formulaString = null;
                                    } else {
                                    }

                                } catch (Exception c) {
                                    //  System.out.println(" Convert data type is false  value=" + o + " , index column: " + jj);
                                }
                            }
                        } catch (Exception x) {
                        }
                    }
                } catch (Exception v) {
                }
            }

            // reset header display
            Set<String> keyset = hm.keySet();
            for (int n = 0; n <= adsColumn + 5; n++) {
                try {
                    if (n == 0) {
                        headerRow.getCell(0).setCellValue(" ");
                    } else {
                        String m = headerRow.getCell(n).getStringCellValue();
                        String headerTemp = "";
                        String nn = "";
                        // boolean contentHeader =false;
                        if (m != null && !("".equals(m))) {
//                           Hashtable<String,String> htm = new Hashtable() ;                           
                            for (String tempKey : keyset) {
                                if (m.contains(tempKey)) {
                                    nn = hm.get(tempKey);
                                    headerTemp = m.replace(tempKey, nn);
                                    // contentHeader=true;
                                    break;
                                }

                            }

                            if (headerTemp != null && headerTemp.length() > 1) {
                                headerRow.getCell(n).setCellValue(headerTemp);   // Tạm cất không thay đổi nội dung tiêu đề.  
                            }
                        }
                    }
                } catch (Exception c) {

                }
            }
            // for loop for set index-(so thu tu trong bang tinh) .
            for (int i = 0; i < adsRows; i++) {
                try {
                    // System.out.println("set value " +i);
                    XSSFRow r = sheet.getRow(i + headerIndex);
                    XSSFCell c = r.createCell(0);
                    if (i == 0) {
                        c.setCellValue(indexHeaderDisplay);
                        continue;
                    }
                    c.setCellValue(i);
                } catch (Exception x) {
                    System.out.println("set value fail  " + i + ", " + x.getMessage());
                }
            }

            cleanSheet(sheet1);
            sheet.removeRow(styleRow);
            //sheet.removeRow(formulaRow);
            //FileOutputStream fos =new FileOutputStream(new File("H:/test/test.xlsx"));
            //workbook.write(fos);

            System.out.println("BEGIN write file Response");
            //  workbook.write(response.getOutputStream());
            //  workbook.close();
            //  in.close();
            System.out.println("FILE  Done Response ");
            return workbook;
        } catch (IOException x) {
            System.out.println("LoI Thao tac file ");
            x.printStackTrace();

        }
        return null;
    }

    static void c() {
        Report r = new Report();

    }
}
