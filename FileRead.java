import java.io.*;
import java.util.ArrayList;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.FileNotFoundException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;

public class FileRead{
    static boolean debugMode = false;
    static ArrayList<String> ID = new ArrayList<String>();
    static ArrayList<String> Name = new ArrayList<String>();
    static ArrayList<String> Tel = new ArrayList<String>();
    static ArrayList<String> Address = new ArrayList<String>();
    static ArrayList<String> Post = new ArrayList<String>();
    static ArrayList<String> Topic = new ArrayList<String>();
    static ArrayList<String> Description = new ArrayList<String>();
    static ArrayList<String> Message = new ArrayList<String>();
    public static void main(String[] args) throws IOException {
        //readHtml();
        String filePath = "data.xlsx";
        //printExcel(filePath);
        fillData(filePath);
        writeHtml();
    }
    static private void fillData(String filePath) throws IOException {
        FileInputStream inputStream = new FileInputStream(new File(filePath));
        
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet1 = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet1.iterator();
        
        int numOfCell = sheet1.getRow(0).getLastCellNum();
        int currRow = 0;
        DataFormatter dataFormatter = new DataFormatter();
        while (rowIterator.hasNext()) {
           Row row = rowIterator.next();
           Iterator<Cell> cellIterator = row.cellIterator();
           
           for(int currCell = 0; currCell < numOfCell ; currCell++) {
               //Cell cell = cellIterator.next();
               Cell cell = sheet1.getRow(currRow).getCell(currCell);
               String cellToString = dataFormatter.formatCellValue(cell);
               if(cell == null) {
                   cellToString = "";
               }
               if(currRow != 0) {
                   switch (currCell) {
                       case 0:
                            ID.add(cellToString);
                            break;
                       case 1:
                            Name.add(cellToString);
                            break;
                       case 2:
                            Tel.add(cellToString);
                            break;
                       case 3:
                            Address.add(cellToString);
                            break;
                       case 4:
                            Post.add(cellToString);
                            break;
                       case 5:
                            Topic.add(cellToString);
                            break;
                       case 6:
                            Description.add(cellToString);
                            break;
                       case 7:
                            Message.add(cellToString);
                            break;     
                       default:
                            break;
                   }
               }
           }
           currRow++;
        }
        workbook.close();
        inputStream.close();
        
        /*
        System.out.println(ID);
        System.out.println(Name);
        System.out.println(Tel);
        System.out.println(Address);
        System.out.println(Post);
        System.out.println(Topic);
        System.out.println(Description);
        System.out.println(Message);
        */
    }
    static private void printTxt(String filePath) throws IOException {
        ArrayList<String> data = new ArrayList<String>();
        File file = new File("hello.txt");
        System.out.println(file.getCanonicalPath());
        FileInputStream ft = new FileInputStream(file);

        DataInputStream in = new DataInputStream(ft);
        BufferedReader br = new BufferedReader(new InputStreamReader(in));
        String strline;            
        while((strline = br.readLine()) != null){
           data.add(strline);
           System.out.println(strline);
        }
        in.close();
    }
    static private void printExcel(String filePath) throws IOException {
        FileInputStream inputStream = new FileInputStream(new File(filePath));
        
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet1 = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet1.iterator();
        
        System.out.println("row: "+ sheet1.getRow(0).getLastCellNum());
        while (rowIterator.hasNext()) {
           Row row = rowIterator.next();
           Iterator<Cell> cellIterator = row.cellIterator();
           
           while (cellIterator.hasNext()) {
               Cell cell = cellIterator.next();
               if(cell != null) {
                   switch (cell.getCellType()) {
                       case STRING:
                            System.out.print(cell.getStringCellValue());
                            break;
                       case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue());
                            break;
                       case NUMERIC:
                            System.out.print(cell.getNumericCellValue());
                            break;
                       case BLANK:
                            System.out.print("BLANK");
                            break;
                       default:
                            break;
                   }
               }
               System.out.print(" - ");
           }
           System.out.println();
        }
        workbook.close();
        inputStream.close();
    }
    static private boolean checkFileExist(String fileName) {
        File checkfile = new File(fileName);
        if (checkfile.exists()){
            return true;
        } else {
            return false;
        }
    }
    static private void readHtml() throws IOException {
        File htmlConfig = new File("config.html");
        BufferedReader br = new BufferedReader(new FileReader(htmlConfig));         
        String st; 
        while ((st = br.readLine()) != null) 
        {
            System.out.println(st.replaceAll("__ID__", "__anutie__")); 
        }
    }
    static private void writeHtml() throws IOException {
        Writer writer = null;
        File htmlConfig = new File("config.html");         
        String st; 
        File directory = new File("output");
        if (! directory.exists()) {
                new File("output").mkdirs();
            }
        for(int i=0;i<ID.size();i++) {
            directory = new File(ID.get(i));
            if (! directory.exists()) {
                new File(ID.get(i)).mkdirs();
            }
            String path = "output/" + ID.get(i) + ".html";
            File outputHtml = new File(path);
            writer = new BufferedWriter(new OutputStreamWriter(
            new FileOutputStream(path), "utf-8"));
            BufferedReader br = new BufferedReader(new FileReader(htmlConfig));
            while ((st = br.readLine()) != null) {
                String outputStr = st;
                outputStr = outputStr.replaceAll("__ID__", ID.get(i));
                outputStr = outputStr.replaceAll("__NAME__", Name.get(i));
                outputStr = outputStr.replaceAll("__TELE__", Tel.get(i));
                outputStr = outputStr.replaceAll("__ADDR__", Address .get(i));
                outputStr = outputStr.replaceAll("__POST__", Post.get(i));
                outputStr = outputStr.replaceAll("__TOPIC__", Topic.get(i));
                outputStr = outputStr.replaceAll("__DESC__", Description.get(i));
                outputStr = outputStr.replaceAll("__MSG__", Message.get(i));
                
                if (checkFileExist(ID.get(i)+"/profile.jpg")) {
                    outputStr = outputStr.replaceAll("__PROFILE_PIC__", "profile.jpg");
                } else if(checkFileExist(ID.get(i)+"/profile.png")) {
                    outputStr = outputStr.replaceAll("__PROFILE_PIC__", "profile.png");
                } else {
                    outputStr = outputStr.replaceAll("__PROFILE_PIC__", "profile.jpg\" style=\"display:none\"");
                }
                
                if (checkFileExist(ID.get(i)+"/QR.jpg")) {
                    outputStr = outputStr.replaceAll("__QR__", "QR.jpg");
                } else if(checkFileExist(ID.get(i)+"/QR.png")) {
                    outputStr = outputStr.replaceAll("__QR__", "QR.png");
                } else {
                    outputStr = outputStr.replaceAll("__QR__", "QR.jpg\" style=\"display:none\"");
                }
                writer.write(outputStr);
            }
            String str = "\u5149\u5fa9\u9999\u6e2f\uff0c\u6642\u4ee3\u9769\u547d\u3002";
            byte[] charset = str.getBytes("UTF-8");
            String lhkroot = new String(charset, "UTF-8");        
            writer.write("<div align=\"center\"><h7>"+lhkroot+"</h7></div>");
            writer.close();
        }
    }
    static private void printDebugMsg(String debugMsg) {
        if(debugMode == true)
            System.out.println(debugMsg);
    }
}