package org.example;
import java.io.*;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.util.*;

import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
//@SuppressWarnings("deprecation")
public class Main {
    public static void main(String[] args) throws IOException,FileNotFoundException {
        Pair<String, String> pair = new Pair<>("","");
        pair.getKey();
        List<Pair<String,String>> list = new ArrayList<>();
        list.add(new Pair<>("",""));

        String path = "/Users/santoshchavva/Downloads/BSNL/revenueshare.xlsx";
        File fileObj = new File(URLDecoder.decode(path, StandardCharsets.UTF_8));
        FileInputStream file=new FileInputStream(fileObj);
        HashMap<String,String> data=new HashMap<>();
        XSSFWorkbook reading=new XSSFWorkbook(file);
        XSSFWorkbook writing=new XSSFWorkbook();
        XSSFSheet readSheet=reading.getSheetAt(0);
        XSSFSheet writesheet=writing.createSheet("generated list");
        Iterator<Row> rowIterator=readSheet.iterator();
       while (rowIterator.hasNext()){
            Row row= rowIterator.next();
            Cell serviceNumber=row.getCell(6);
            Cell revenueShare=row.getCell(28);
            String key=getCellValue(serviceNumber);
            String value=getCellValue(revenueShare);
            if (key!=null && value!=null){
                data.put(key,value);

            }
        }
        Set<String> keyset=data.keySet();
        int rownum=0;
        for (Map.Entry<String, String> entry : data.entrySet()) {
            Row row = writesheet.createRow(rownum++);
            Cell keyCell = row.createCell(0);
            keyCell.setCellValue(entry.getKey());
            Cell valueCell = row.createCell(1);
            valueCell.setCellValue(entry.getValue());
        }
        FileOutputStream fileOutputStream=new FileOutputStream(new File("/Users/santoshchavva/Downloads/BSNL/generated list.xlsx"));
        writing.write(fileOutputStream);
        fileOutputStream.close();
        file.close();
        data.forEach((k, v) -> System.out.println("Key: " + k + ", Value: " + v));
        System.out.println(data.size());
    }

    private static String getCellValue(Cell cell) {
        if(cell.getCellType()== CellType.NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        }
        else if (cell.getCellType()==CellType.STRING){
            return cell.getStringCellValue();
        }
        else{
            return null;

        }
    }
}
