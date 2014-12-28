package ja;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;

import javax.swing.JFileChooser;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class App {

	public static void main(String[] args) {
		
		JFileChooser fc = new JFileChooser();
		fc.setFileSelectionMode( JFileChooser.DIRECTORIES_ONLY );
		fc.showOpenDialog(null);
		
		File sourceDir = fc.getSelectedFile();
		
		if( sourceDir != null)
		{
			//File sourceDir = new File( "F:/Projects/Programming/TestData/BavaDataMatrix2Column/" );
			File[] fileList = sourceDir.listFiles( );
			
			Workbook wb = new HSSFWorkbook();
	   	    FileOutputStream fileOut;
	   	    
	   	    Sheet sheet1 = wb.createSheet( "data" );
	   	    
	   	    int fileCount = -1;    // used to keep track of what column we are in.
			
	   	    try 
	   	    {
	   	    	
				for( int i = 0; i < fileList.length; i++)
				{
					String curFileName = fileList[ i ].getName();
					
					// Does this file end in jnd
					if( curFileName.lastIndexOf('.') > 0 )
		            {
						
						// get last index for '.' char
						int lastIndex = curFileName.lastIndexOf('.');
		               
						// get extension
						String extension = curFileName.substring(lastIndex);
		               
						// match path name extension
						if(extension.equals(".jnd"))
						{
							fileCount++;  // increment the file (column) we are on.
							
							String filenameBase = curFileName.substring(0, lastIndex);
							
							Row topRow = sheet1.getRow( 0 );
							
							if( topRow == null )
	           				{
								topRow = sheet1.createRow( 0 );
	           				}
							
							Cell curCell = topRow.createCell( fileCount );
	           				
	           				curCell.setCellValue( filenameBase );
							
							System.out.println( filenameBase );
							
							String xlsFilename = filenameBase + ".xls";
							
		            	   	
			           			FileInputStream inFile = new FileInputStream( fileList[i] );
			           			InputStreamReader inFile2 = new InputStreamReader( inFile );
			           			BufferedReader inFile3 = new BufferedReader( inFile2 );
			           			String inLine;
			           			
			           			int lineCount = 0;
			           			
			           			while( ( inLine = inFile3.readLine() ) != null )
			           			{
			           				System.out.println( inLine );
			           				lineCount++;
			           			}
			           			
			           			inFile3.close();
			           			
			           			inFile = new FileInputStream( fileList[i] );
			           			inFile2 = new InputStreamReader( inFile );
			           			inFile3 = new BufferedReader( inFile2 );
			           			
			           			int rowCount = 1;
			           			
			           			for( int j = 0; j < lineCount - 1; j++)
			           			{
			           				inLine = inFile3.readLine();
			           				
			           				
			           				
			           				for( int k = 0; k < (lineCount - 1) - j; k++)
			           				{
				           				String readData = inLine.substring( 30 + (k * 15) + (j * 15), 30 + (k * 15) + (j * 15) + 6);
				           				System.out.println( readData );
				           				
				           				Row curRow = sheet1.getRow( rowCount );
				           				
				           				if( curRow == null )
				           				{
				           					curRow = sheet1.createRow( rowCount );
				           				}
				           				
				           				curCell = curRow.createCell( fileCount );
				           				
				           				curCell.setCellValue( readData );
				           				
				           				rowCount++;
			           				}
			           			}
			           			
			           			
			           			
			           			inFile3.close();
			           			
			           		
						}
		            }
					
					
				}
			
				fileOut = new FileOutputStream( sourceDir + "/data.xls" );// + xlsFilename);
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

}
