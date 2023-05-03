package plakamaven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.Queue;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.Map;
import java.util.Arrays;
import java.util.Calendar;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

public class App 
{

    public static void main( String[] args ) throws IOException, InvalidFormatException
    {
        File doc = new File("/home/mitsos/Desktop/plaka1.xls");//the excell downloaded from bookings extranet
        FileInputStream inputStream = new FileInputStream(doc);
        Workbook inputExcel = new HSSFWorkbook(inputStream);
        Sheet sheet = inputExcel.getSheetAt(0);
        doc = new File("/home/mitsos/Desktop/plaka2.xls");//the excell copied and pasted to a new excel from booking extranet
        FileInputStream inputStreamRooms = new FileInputStream(doc);
        Workbook inputExcelRooms = new HSSFWorkbook(inputStreamRooms);
        Sheet sheetForRoomsOnly = inputExcelRooms.getSheetAt(0);           
        Map<String, Queue<String>> values = new HashMap<>();
        HashMapCreator(values);
        fromEcxelToHashMap(values, sheet, sheetForRoomsOnly);
        inputExcel.close();
        inputExcelRooms.close();
        Workbook outputExcel = fromHashMapToExcelFNL(values);
        FileOutputStream outputStream = new FileOutputStream("output.xlsx");
        outputExcel.write(outputStream);
        outputStream.close();
    }

    public static void HashMapCreator(Map<String, Queue<String>> values){
        Queue<String> checkInValues = new LinkedList<String>();
        Queue<String> checkOutValues = new LinkedList<String>();
        Queue<String> statusValues = new LinkedList<String>();
        Queue<String> roomsValues = new LinkedList<String>();
        Queue<String> peopleValues = new LinkedList<String>();
        Queue<String> childrenValues = new LinkedList<String>();
        Queue<String> remarksValues = new LinkedList<String>();
        values.put("checkIn", checkInValues);
        values.put("checkOut", checkOutValues);
        values.put("status", statusValues);
        values.put("rooms", roomsValues);
        values.put("people", peopleValues );
        values.put("children", childrenValues);
        values.put("remarks", remarksValues);
    }

    public static void fromEcxelToHashMap(Map<String, Queue<String>> values, Sheet sheet, Sheet sheetForRoomsOnly){
        //below code to manipulate sheet and add checkIn, checkOut, status, remarks, people and children to hashmap
        Row row;
        Cell cell;
        Iterator<Row> rowIterator;
        for(int colIndex = 0; colIndex < sheet.getRow(0).getLastCellNum(); colIndex++) {
            if(colIndex == 3 || colIndex == 4 || colIndex == 6 || colIndex == 7 || colIndex == 8 || colIndex == 10 || colIndex == 17){
                int i = 0;
                rowIterator = sheet.rowIterator();
                while(rowIterator.hasNext()){
                    row = sheet.getRow(i);
                    cell = row.getCell(colIndex);
                    if (cell.getCellType() == CellType.STRING){
                        //i have to exclude i = 0, so to not save the headers value, sinxe the iterator starts at row 0 by default
                        //which means that the cells of the first row will be iterated no matter what
                        if(i!=0){
                            if(colIndex == 3){
                                values.get("checkIn").add(cell.getStringCellValue());   
                            }
                            if(colIndex == 4){                                  
                                values.get("checkOut").add(cell.getStringCellValue());   
                            }
                            if(colIndex == 6){                                  
                                values.get("status").add(cell.getStringCellValue());   
                            }
                            if(colIndex == 17){    
                                String parsedRemarks = cell.getStringCellValue();
                                Pattern pattern = Pattern.compile("(?<=arrival: between )(\\d{2}:\\d{2}) and (\\d{2}:\\d{2})");
                                Matcher matcher = pattern.matcher(parsedRemarks);  
                                if(matcher.find()){
                                    String match = matcher.group(1);   
                                    values.get("remarks").add(match);   
                                }
                                else  values.get("remarks").add("empty");    
                            }
                        }
                    }
                    else if (cell.getCellType() == CellType.NUMERIC){
                        //no i = someval needs to be excluded beacuse the first row(headers) only has cells of type STRING
                        //so this if statement only reads values from NUMERIC type cells
                         if(colIndex == 7){
                            /*values.get("rooms").add(Double.toString(cell.getNumericCellValue()));
                            auto tha douleue teleia gamo ton xristo mou an den gamiotan to arxeio 
                            otan allazo ta domatia manually apo to excel. ant autou bazv manually ta
                            domatia stis grammes 129-136*/
                        } 
                        if(colIndex == 8){
                            values.get("people").add(Double.toString(cell.getNumericCellValue()));
                        } 
                        if(colIndex == 10){
                            values.get("children").add(Double.toString(cell.getNumericCellValue()));     
                        }
                    }
                    else if (cell.getCellType() == CellType.BLANK){
                        //no i = someval needed for the same reaso as NUMERIC above
                        if(colIndex == 17){                                  
                            values.get("remarks").add("empty");   
                        }
                        if(colIndex == 10){
                            values.get("children").add("0");
                        }
                    }
                    row = rowIterator.next(); 
                    i++;                         
                }
            }
        }
        //below code to manipulate sheetForRoomsOnly to add rooms to hashmap since the data is collected from another file
        rowIterator = sheetForRoomsOnly.rowIterator();
        while(rowIterator.hasNext()){
            row = rowIterator.next();
            cell = row.getCell(3);
            if(cell != null && cell.getCellType() == CellType.NUMERIC){
                values.get("rooms").add(Double.toString(cell.getNumericCellValue()).replaceAll(".0", ""));
            }
            else if(cell != null && cell.getCellType() == CellType.STRING){
                values.get("rooms").add(cell.getStringCellValue().replaceAll("1 x", "").replaceAll("[^\\d]", ""));
            }
        }
    }

    public static XSSFWorkbook fromHashMapToExcelFNL(Map<String, Queue<String>> values){
        XSSFWorkbook outputEXcel = sheetCreator(values);
        outputEXcel = sheetFiller(outputEXcel, values);
        return outputEXcel;
    }

    private static XSSFWorkbook sheetCreator(Map<String, Queue<String>> values){
        XSSFWorkbook outputEXcel = new XSSFWorkbook();
        XSSFSheet sheet = outputEXcel.createSheet("Sheet1");
        String stringMonth = "" + (values.get("checkIn").peek()).charAt(5) + (values.get("checkIn").peek()).charAt(6);
        String stringYear= "" + (values.get("checkIn").peek()).charAt(0) + (values.get("checkIn").peek()).charAt(1) + (values.get("checkIn").peek()).charAt(2) + (values.get("checkIn").peek()).charAt(3);
        int year = Integer.parseInt(stringYear);
        int month = Integer.parseInt(stringMonth);
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.YEAR, year);
        calendar.set(Calendar.MONTH, month - 1); // January is month 0
        int daysOfMonth = calendar.getActualMaximum(Calendar.DAY_OF_MONTH);
        //matrix Creation
        for (int row = 0; row < daysOfMonth + 1; row++) {
            Row sheetRow = sheet.createRow(row);
            for (int col = 0; col < 22; col++) {
                if(col == 1 || col == 20){
                    Cell cell = sheetRow.createCell(col, CellType.NUMERIC);//to double einai primitive kai den pairnei null value
                    cell.setCellValue(0); 
                }
                else{
                    Cell cell = sheetRow.createCell(col, CellType.STRING);//epidhd to String einai reference type kai oxi
                    cell.setCellValue((String)null);                      //primitive mporo na valo null value
                }
            }
        }
        //putting the corect headers
        String stringHeaderTags[] = {calendar.getDisplayName(Calendar.MONTH, Calendar.LONG, java.util.Locale.getDefault()) , "Deluxe" , "1" , " " ,  "Exceptional", "2" , " " , "Executive" , "3" , " ", "Superior" , "4" , " " , "Classic" , "5" , " " , "Standard" , "6" , " ", calendar.getDisplayName(Calendar.MONTH, Calendar.LONG, java.util.Locale.getDefault())};
        Queue<String> tags = new LinkedList<>(Arrays.asList(stringHeaderTags));
        for (int col = 0; col < 22; col++) {
            if(col == 1 || col == 20){
                //edo bazo tous arithmous ton hmeron dipla sta onomata stous
                if(col == 1 ||col == 20){
                    for(int row = 1; row < daysOfMonth + 1; row++){
                        sheet.getRow(row).getCell(col).setCellValue(row);
                    }
                }
            }
            else{
                //edp bazo ta headers
                String val = tags.poll();
                sheet.getRow(0).getCell(col).setCellValue(val);
                //edo bazo tis sostes meres me taonomata tous sthn proth kai teleutaia sthlh
                if(col == 0 ||col == 21){
                    for(int row = 1; row < daysOfMonth + 1; row++){
                        calendar.set(Calendar.DAY_OF_MONTH, row);
                        String nameOfDay = calendar.getDisplayName(Calendar.DAY_OF_WEEK, Calendar.LONG, java.util.Locale.getDefault());
                        sheet.getRow(row).getCell(col).setCellValue(nameOfDay);
                    }
                }
            }            
        }        
        return outputEXcel;
    }

    private static XSSFWorkbook sheetFiller(XSSFWorkbook outputEXcel, Map<String, Queue<String>> values){
        XSSFSheet sheet = outputEXcel.getSheetAt(0);
        Map<Integer, Integer> roomsMathematic = new HashMap<Integer, Integer>();
        roomsMathematic.put(1,2);
        roomsMathematic.put(2,5);
        roomsMathematic.put(3,8);
        roomsMathematic.put(4,11);
        roomsMathematic.put(5,14);
        roomsMathematic.put(6,17);
        while(!values.get("checkIn").isEmpty()){
            //an to status den einai "ok" diagrafo thn sygkekrimenh krathsh apo to hashmap
            if(!values.get("status").peek().equalsIgnoreCase("ok")){
                values.get("status").remove();
                values.get("checkIn").remove();
                values.get("checkOut").remove();
                values.get("rooms").remove();
                values.get("people").remove();
                values.get("children").remove();
                values.get("remarks").remove();
            }
            //allios mpaino na epeskergasto thn sygkekrimenh krathsh
            else{
                String temp = values.get("rooms").remove();//room of book removed;
                int room = Integer.parseInt(temp);
                values.get("status").remove();  //status of book removed.(status "ok" afou mphka edo kai prepei na to diagrapso tora)
                temp = "" + values.get("checkIn").peek().charAt(5) + values.get("checkIn").peek().charAt(6) + values.get("checkIn").peek().charAt(8) + values.get("checkIn").peek().charAt(9);
                values.get("checkIn").remove();     //checkin of book removed
                int startOfBook = Integer.parseInt(temp);
                temp = "" + values.get("checkOut").peek().charAt(5) + values.get("checkOut").peek().charAt(6) + values.get("checkOut").peek().charAt(8) + values.get("checkOut").peek().charAt(9);
                values.get("checkOut").remove();    //checkout of book removed
                int endOfBook = Integer.parseInt(temp);
                //parakato boolean gia krathseis pou exoun checkout se allo mhna.to pernao san parametro se kathe handling case parakato
                //false an checkin checkout einai ston idio mhna allios true
                boolean continuesToNextMonth = false;
                if(startOfBook/100 != endOfBook/100){
                    continuesToNextMonth = true;
                }           
                //handling gia book monhs krathsh (dld ena mono domatio)
                if(room/10 == 0){
                    collBookPeopleCheckOut(startOfBook, endOfBook, sheet, values, room, continuesToNextMonth, roomsMathematic);
                    //people, children and remarks of book of are getting removed in the function call
                }
                //handling gia pollaples krathseis (dld h krathsh na einai gid 2 kai 3 diamerismata, mexri kai 6 ola dld)
                else{
                    int multipleRooms = String.valueOf(room).length();
                    int primaryRoom;
                    String secondaryRooms = Integer.toString(room);
                    switch(multipleRooms){
                        case 2://kathe case afora ton arithmo domation sthn sygkekrimenh krathsh dld case 2, posothta dyo domation, case 3 trion klp
                            primaryRoom = room/10;
                            collBookPeopleCheckOut(startOfBook, endOfBook, sheet, values, primaryRoom, continuesToNextMonth, roomsMathematic);
                            collAssosiatedBook(startOfBook, endOfBook, sheet, values, Character.getNumericValue(secondaryRooms.charAt(1)), primaryRoom, continuesToNextMonth, roomsMathematic);
                        break;
                        case 3:
                            primaryRoom = room/100;
                            collBookPeopleCheckOut(startOfBook, endOfBook, sheet, values, primaryRoom, continuesToNextMonth, roomsMathematic);   
                            for(int i = 1; i < 3; i++){
                                collAssosiatedBook(startOfBook, endOfBook, sheet, values, Character.getNumericValue(secondaryRooms.charAt(i)), primaryRoom, continuesToNextMonth, roomsMathematic);
                            }                          
                        break;
                        case 4:
                            primaryRoom = room/1000;
                            collBookPeopleCheckOut(startOfBook, endOfBook, sheet, values, primaryRoom, continuesToNextMonth, roomsMathematic);
                            for(int i = 1; i < 4; i++){
                                collAssosiatedBook(startOfBook, endOfBook, sheet, values, Character.getNumericValue(secondaryRooms.charAt(i)), primaryRoom, continuesToNextMonth, roomsMathematic);
                            }   
                        break;
                        case 5:
                            primaryRoom = room/10000;
                            collBookPeopleCheckOut(startOfBook, endOfBook, sheet, values, primaryRoom, continuesToNextMonth, roomsMathematic);
                            for(int i = 1; i < 5; i++){
                                collAssosiatedBook(startOfBook, endOfBook, sheet, values, Character.getNumericValue(secondaryRooms.charAt(i)), primaryRoom, continuesToNextMonth, roomsMathematic);
                            }   
                        break;
                        case 6:
                            primaryRoom = room/100000;
                            collBookPeopleCheckOut(startOfBook, endOfBook, sheet, values, primaryRoom, continuesToNextMonth, roomsMathematic);
                            for(int i = 1; i < 6; i++){
                                collAssosiatedBook(startOfBook, endOfBook, sheet, values, Character.getNumericValue(secondaryRooms.charAt(i)), primaryRoom, continuesToNextMonth, roomsMathematic);
                            }   
                        break; 
                    }
                }                    
                
            }
        }
        return outputEXcel;
    }

    private static void collBookPeopleCheckOut(int startOfBook, int endOfBook, XSSFSheet sheet, Map<String, Queue<String>> values, int room, boolean continuesToNextMonth, Map<Integer, Integer> roomsMathematic){
        int a = roomsMathematic.get(room);
        if(!continuesToNextMonth){
            for(int i = startOfBook%100; i < endOfBook%100; i ++){
                if(i == startOfBook%100){
                    sheet.getRow(i).getCell(a).setCellValue("startOfBook");
                }
                else sheet.getRow(i).getCell(a).setCellValue("Book");
                double people = Double.parseDouble(values.get("people").peek());
                Double children = Double.parseDouble(values.get("children").peek());
                Double adultsAndChildren = people - children;
                children /= 10;
                adultsAndChildren += children;
                if(Double.toString(adultsAndChildren).contains("0")){
                    sheet.getRow(i).getCell(a + 1).setCellValue(Integer.toString(Double.valueOf(adultsAndChildren).intValue()));
                }
                else  sheet.getRow(i).getCell(a + 1).setCellValue(Double.toString(adultsAndChildren));
            }
            if(values.get("remarks").peek() != "empty"){
                sheet.getRow(startOfBook%100).getCell(a + 2).setCellValue(values.get("remarks").peek());
            }
        }
        //handling book for single room book that check in and check out are NOT in the same month
        else{
            int endOfMonth = sheet.getLastRowNum() + 1;
            if(startOfBook%100 == endOfMonth - 1){
                sheet.getRow(startOfBook%100).getCell(a).setCellValue("SOB(continues(" + endOfBook%100 + "/" + endOfBook/100  + ")");
                double people = Double.parseDouble(values.get("people").peek());
                Double children = Double.parseDouble(values.get("children").peek());
                Double adultsAndChildren = people - children;
                children /= 10;
                adultsAndChildren += children;
                sheet.getRow(startOfBook%100).getCell(a + 1).setCellValue(Double.toString(adultsAndChildren));
            }
            else{
                for(int i = startOfBook%100; i < endOfMonth; i ++){
                    if(i == startOfBook%100){
                        sheet.getRow(i).getCell(a).setCellValue("startOfBook");
                    }
                    else if(i == endOfMonth - 1){
                        sheet.getRow(i).getCell(a).setCellValue("continues(" + endOfBook%100 + "/" + endOfBook/100  + ")");
                    }
                    else sheet.getRow(i).getCell(a).setCellValue("Book");
                    double people = Double.parseDouble(values.get("people").peek());
                    Double children = Double.parseDouble(values.get("children").peek());
                    Double adultsAndChildren = people - children;
                    children /= 10;
                    adultsAndChildren += children;
                    if(Double.toString(adultsAndChildren).contains("0")){
                        sheet.getRow(i).getCell(a + 1).setCellValue(Integer.toString(Double.valueOf(adultsAndChildren).intValue()));
                    }
                    else  sheet.getRow(i).getCell(a + 1).setCellValue(Double.toString(adultsAndChildren));
                }
            }
            if(values.get("remarks").peek() != "empty"){
                sheet.getRow(startOfBook%100).getCell(a + 2).setCellValue(values.get("remarks").peek());
            }

        }
        values.get("people").remove();     
        values.get("children").remove();   
        values.get("remarks").remove();   
    }

    //helper function that works togther with the other two for multiple rooms in a single booking
    private static void collAssosiatedBook(int startOfBook, int endOfBook, XSSFSheet sheet, Map<String, Queue<String>> values, int secondaryRoom, int primaryRoom, boolean continuesToNextMonth, Map<Integer, Integer> roomsMathematic){
        //handling check in and check out are in the same month  
        int a = roomsMathematic.get(secondaryRoom);   
        if(!continuesToNextMonth){
            for(int i = startOfBook%100; i < endOfBook%100; i ++){
                //first col = book
                if(i == startOfBook%100){
                    sheet.getRow(i).getCell(a).setCellValue("GroupRoom" + primaryRoom);
                    if(sheet.getRow(i).getCell(roomsMathematic.get(primaryRoom)).getStringCellValue() != "GroupRoom" + Integer.toString(primaryRoom)){
                        sheet.getRow(i).getCell(roomsMathematic.get(primaryRoom)).setCellValue("GroupRoom" + primaryRoom);
                    }
                }
                else sheet.getRow(i).getCell(a).setCellValue("Book");
            }
        }
        //handling check in and check out are NOT in the same month    
        else{
            int endOfMonth = sheet.getLastRowNum() + 1;
            if(startOfBook%100 == endOfMonth - 1){
                sheet.getRow(endOfMonth - 1).getCell(a).setCellValue("GroupRoom" + primaryRoom + "continues(" + endOfBook%100 + "/" + endOfBook/100  + ")");
                sheet.getRow(endOfMonth - 1).getCell(roomsMathematic.get(primaryRoom)).setCellValue("GroupRoom" + primaryRoom + "continues(" + endOfBook%100 + "/" + endOfBook/100  + ")");
                
            }
            else{
                for(int i = startOfBook%100; i < endOfMonth; i ++){
                    if(i == startOfBook%100){
                        sheet.getRow(i).getCell(a).setCellValue("GroupRoom" + primaryRoom);
                        if(sheet.getRow(i).getCell(roomsMathematic.get(primaryRoom)).getStringCellValue() != "GroupRoom" + Integer.toString(primaryRoom)){
                            sheet.getRow(i).getCell(roomsMathematic.get(primaryRoom)).setCellValue("GroupRoom" + primaryRoom);
                        }
                    }
                    else if(i == endOfMonth - 1){
                        sheet.getRow(i).getCell(a).setCellValue("continues(" + endOfBook%100 + "/" + endOfBook/100  + ")");
                    }
                    else sheet.getRow(i).getCell(a).setCellValue("Book");
                }
            }
        }
    }
}
