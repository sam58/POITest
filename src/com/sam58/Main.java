package com.sam58;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    private static Logger logger = Logger.getLogger(Main.class);

    public static void main(String[] args) throws IOException {
        logger.warn("start");
        try {
            File excel = new File(args[0]);
            FileInputStream fis = new FileInputStream(excel);

            XSSFWorkbook book = new XSSFWorkbook(fis);
            XSSFSheet sheet = book.getSheetAt(0);

            Iterator<Row> itr = sheet.iterator();

            // Iterating over Excel file in Java
            while (itr.hasNext()) {
                Row row = itr.next();

                // Iterating over each column of Excel file
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();

                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;
                        default:

                    }
                }
                System.out.println("");
            }
            logger.warn("begin generation");
            // writing data into XLSX file
            Map<String, Object[]> newData = new HashMap<String, Object[]>();
           // for(Long gg=0L;gg<10000L;gg++) {
           //     newData.put(gg.toString(), new Object[]{gg.toString(), "Sonya", "75K--"+gg.toString(), "SALES", "Rupert"});
           // }
            Set<String> newRows = newData.keySet();
            int rownum = sheet.getLastRowNum();
            Iterator<String> its= newRows.iterator();

            //for (String key : newRows) {
            for(Long gg=0L;gg<1000L;gg++) {//строк

               // {   new Object[]{gg.toString(), "Sonya", "75K--"+gg.toString(), "SALES", "Rupert"});
                Row row = sheet.createRow(rownum++);
               // Object[] objArr = newData.get(key);
                Object[] objArr= new Object[280] ;
                for (int jj=0;jj<280;jj++){//столбцов
                    objArr[jj]= gg.toString()+"ddd"+jj;
                }
                int cellnum = 0;
                for (Object obj : objArr) {
                    Cell cell = row.createCell(cellnum++);
                    if (obj instanceof String) {
                        cell.setCellValue((String) obj);
                    } else if (obj instanceof Boolean) {
                        cell.setCellValue((Boolean) obj);
                    } else if (obj instanceof Date) {
                        cell.setCellValue((Date) obj);
                    } else if (obj instanceof Double) {
                        cell.setCellValue((Double) obj);
                    }
                }
            }

            // open an OutputStream to save written data into Excel file
            FileOutputStream os = new FileOutputStream(args[1]);
            book.write(os);
            System.out.println("Writing on Excel file Finished ...");

            // Close workbook, OutputStream and Excel file to prevent leak
            os.close();

            fis.close();

        } catch (FileNotFoundException fe) {
            fe.printStackTrace();
        } catch (IOException ie) {
            ie.printStackTrace();
        }catch (Exception e ){
            e.printStackTrace();
        }
       logger.warn("end");
    }
}
