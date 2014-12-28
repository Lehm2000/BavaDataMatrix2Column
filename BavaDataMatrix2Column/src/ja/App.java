package ja;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class App {

	public static void main(String[] args) {

		System.out.println( "yo!");
		Workbook wb = new HSSFWorkbook();
	    FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream("F:/Projects/Programming/TestData/BavaDataMatrix2Column/workbook.xls");
			wb.write(fileOut);
		    fileOut.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	    

	}

}
