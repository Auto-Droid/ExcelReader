import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelReader {
 
	static ArrayList<String> strValues;
	static ArrayList<String> alnumValues;
	static ArrayList<Integer> intValues;
	
	public static void main(String[] arg){
		strValues = new ArrayList<String>();
		intValues = new ArrayList<Integer>();
		alnumValues = new ArrayList<String>();
		readFile();
		 WriteExcelFile();
	}
	
	
	public static void readFile(){
	 File myFile = new File("C:/Book1.xlsx");
     FileInputStream fis = null;
	try {
		fis = new FileInputStream(myFile);
	} catch (FileNotFoundException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}

     // Finds the workbook instance for XLSX file
     XSSFWorkbook myWorkBook = null;
	try {
		myWorkBook = new XSSFWorkbook (fis);
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
    
     // Return first sheet from the XLSX workbook
     XSSFSheet mySheet = myWorkBook.getSheetAt(0);
    
     // Get iterator to all the rows in current sheet
     Iterator<Row> rowIterator = mySheet.iterator();
    
     // Traversing over each row of XLSX file
     while (rowIterator.hasNext()) {
         Row row = rowIterator.next();

         // For each row, iterate through each columns
         Iterator<Cell> cellIterator = row.cellIterator();
         while (cellIterator.hasNext()) {

             Cell cell = cellIterator.next();

             switch (cell.getCellType()) {
             case Cell.CELL_TYPE_STRING:
                 sortStringFromCell(cell.getRichStringCellValue()+"");
                 break;
             case Cell.CELL_TYPE_NUMERIC:
                 System.out.print(cell.getNumericCellValue() + " :Numerical \t");
                 intValues.add((int) cell.getNumericCellValue());
                 break;
             default :
          
             }
         }
         System.out.println("");
     }
 }
	/**
	 * sorts alphanumeric value and string value
	 * @param str
	 */
	public static void sortStringFromCell(String str){
		if(!str.equals("")){
		if(str.matches(".*\\d+.*")){
			alnumValues.add(str);
			System.out.print(str);
		}else{
			strValues.add(str);
			System.out.print(str);
		}
		}
	}
	
	public static void WriteExcelFile(){
		 HSSFWorkbook workbook = new HSSFWorkbook();
		    HSSFSheet sheet = workbook.createSheet("sheet");
		   
		    contentWriter(sheet,0,intValues);
		    contentWriter(sheet,1,alnumValues);
		    contentWriter(sheet,2,strValues);
		    
		    try {
				FileOutputStream out = 
						new FileOutputStream(new File("C:\\new.xls"));
				workbook.write(out);
				out.close();
				System.out.println("Excel written successfully..");
				
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
	}
	
	public synchronized static void contentWriter(HSSFSheet sheet,int count,ArrayList strValues2){
		 //Create a new row in current sheet
	   // Row row = sheet.createRow(count);
	    //Create a new cell in current row
	    int columnNum=count;
	  
	       /*Cell cell1=row.createCell(columnNum);
	        cell1.setCellValue("Numerical");*/
	        //System.out.println(map.get(key));
	        List<Object> columnValues = strValues2;
	        int tempHeight=columnValues.size();
	       
	        int temp=1;
	       for(int i=0;i<strValues2.size();i++)
	        {
	            Row row2;
	            //System.out.println("no of rows:"+(sheet.getPhysicalNumberOfRows()-1)+", height:"+tempHeight);
	            if(sheet.getPhysicalNumberOfRows()-1>temp-1)
	            {
	                //System.out.println("take row");
	                row2=sheet.getRow(temp);

	            }
	            else
	            {

	                //System.out.println("Row inserted");
	                row2=sheet.createRow(temp);
	            }
	            Cell cell2=row2.createCell(columnNum);
	            cell2.setCellValue(""+strValues2.get(i));
	            temp=temp+1;
	        }
	       
	        columnNum=columnNum+1;
	    
	}
	
}
