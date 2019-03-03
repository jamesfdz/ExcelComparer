package compare.excel;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class runComparer {
	
	public runComparer() {}

	
	//method to compare all excel file and excel sheets within them with each other
	public void compare(File[] excelSheetsFilePath) throws IOException {
		
//		XSSFWorkbook final_workbook = new XSSFWorkbook(new FileInputStream("compared_output.xlsx"));
		
		
		XSSFWorkbook book = new XSSFWorkbook(new FileInputStream("processing.xlsx"));
				
		for(File filePath : excelSheetsFilePath) {
			
			FileInputStream fIP = new FileInputStream(filePath);
			XSSFWorkbook b = new XSSFWorkbook(fIP);
			
			System.out.println(filePath+": "+b.getNumberOfSheets());
			
			for (int i = 0; i < b.getNumberOfSheets(); i++) {
				System.out.println(filePath+": "+b.getSheetName(i));
				XSSFSheet sheet = book.createSheet(b.getSheetName(i));
				copySheets(book, sheet, b.getSheetAt(i));
			}
		}
		
		XSSFSheet sheet_1 = book.getSheetAt(0);
		
		if(sheet_1 != null) {
			int index = book.getSheetIndex(sheet_1);
			book.removeSheetAt(index);
		}
		
		FileOutputStream out = new FileOutputStream(new File("compared_result.xlsx"));
		
	    //write operation workbook using file out object
		book.write(out);
	    out.close();
	    book.close();
	    
	    System.out.println("compared_result.xlsx written successfully");
	    
	    System.out.println("Starting to compare the sheets");
	    
	    XSSFWorkbook workbook_compare = new XSSFWorkbook(new FileInputStream("compared_result.xlsx"));
	    
	    XSSFCellStyle style = workbook_compare.createCellStyle();
	    XSSFColor my_background=new XSSFColor(Color.GREEN);
	    style.setFillForegroundColor(my_background);
	    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    
	    XSSFSheet mainSheet = workbook_compare.getSheetAt(0);
	    XSSFSheet checkSheet = null;
	    
	    for(int z = 1; z < workbook_compare.getNumberOfSheets(); z++) {
	    	checkSheet = workbook_compare.getSheetAt(z);
	    	compareDataInBothSheets(mainSheet, checkSheet, style);
	    }
	    
	    System.out.println("Compare finish");
	    
	    System.out.println("Starting to save changes in compared_result.xlsx");
	    
	    FileOutputStream final_out = new FileOutputStream(new File("compared_result.xlsx"));
	    workbook_compare.write(final_out);
	    final_out.close();
	    workbook_compare.close();
	    
	    JOptionPane.showMessageDialog(null, "Completed Successfully");
	    
	}

	
	private void compareDataInBothSheets(XSSFSheet srcSheetMain, XSSFSheet destSheetToCompare, CellStyle style) throws IOException {
		
		int maxRowNumberSrc = srcSheetMain.getLastRowNum() + 1;	
				
		for(int y = srcSheetMain.getFirstRowNum(); y <= maxRowNumberSrc; y++) {
			//getting rows from src and destination
			XSSFRow srcRowMain = srcSheetMain.getRow(y);
						
			if(srcRowMain != null) {
				int firstCellNum = srcRowMain.getFirstCellNum();
				int lastCellNum = srcRowMain.getLastCellNum();
				
				if(lastCellNum != -1) {
					for(int c = firstCellNum; c <= lastCellNum; c++) {
						XSSFCell srcCell = srcRowMain.getCell(c);
						if(srcCell != null && srcCell.getStringCellValue() != "") {
							String srcCellContent = srcCell.getStringCellValue();
							checkContentInDestSheets(srcCellContent, destSheetToCompare, srcSheetMain, style);
						}						
					}
				}
			}
			
		}
	}


	private void checkContentInDestSheets(String srcCellContent, XSSFSheet destSheetToCompare, XSSFSheet srcSheetMain, CellStyle style) throws FileNotFoundException, IOException {
		int maxRowNumberDest = destSheetToCompare.getLastRowNum()+1;
		
		for(int t = destSheetToCompare.getFirstRowNum(); t <= maxRowNumberDest; t++) {
			XSSFRow destRowToCompare = destSheetToCompare.getRow(t);
			
			if(destRowToCompare != null) {
				int firstCellDest = destRowToCompare.getFirstCellNum();
				int lastCellNumDest = destRowToCompare.getLastCellNum();
				
				if(lastCellNumDest != -1) {
					for(int b = firstCellDest; b <= lastCellNumDest; b++) {						
						XSSFCell destCell = destRowToCompare.getCell(b);
						if(destCell != null) {
							String destCellContent = destCell.getStringCellValue();
							if(srcCellContent.equals(destCellContent)) {
								destCell.setCellStyle(style);
							}					
						}
					}
				}
			}
		}
		
	}


	private static void copySheets(XSSFWorkbook book, XSSFSheet sheet, XSSFSheet xssfSheet) {
		copySheets(book, sheet, xssfSheet, true);
	}


	private static void copySheets(XSSFWorkbook book, XSSFSheet sheet, XSSFSheet xssfSheet, boolean copyStyle) {
		int newRownumber = sheet.getLastRowNum() + 1;
		int maxColumnNum = 0; 
		Map<Integer, XSSFCellStyle> styleMap = (copyStyle) ? new HashMap<Integer, XSSFCellStyle>() : null;
		
		for (int i = xssfSheet.getFirstRowNum(); i <= xssfSheet.getLastRowNum(); i++) {     
	      XSSFRow srcRow = xssfSheet.getRow(i);     
	      XSSFRow destRow = sheet.createRow(i + newRownumber);     
	      if (srcRow != null) {     
	        copyRow(book, xssfSheet, sheet, srcRow, destRow, styleMap);     
	        if (srcRow.getLastCellNum() > maxColumnNum) {     
	            maxColumnNum = srcRow.getLastCellNum();     
	        }     
	      }     
	    }
		
		for (int i = 0; i <= maxColumnNum; i++) {     
	      sheet.setColumnWidth(i, xssfSheet.getColumnWidth(i));     
	    }
		
	}


	private static void copyRow(XSSFWorkbook book, XSSFSheet xssfSheet, XSSFSheet sheet, XSSFRow srcRow,
			XSSFRow destRow, Map<Integer, XSSFCellStyle> styleMap) {
		destRow.setHeight(srcRow.getHeight());
		for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {     
	      XSSFCell oldCell = srcRow.getCell(j);
	      XSSFCell newCell = destRow.getCell(j);
	      if (oldCell != null) {     
	        if (newCell == null) {     
	          newCell = destRow.createCell(j);     
	        }     
	        copyCell(book, oldCell, newCell, styleMap);
	      }     
	    }
	}


	@SuppressWarnings("deprecation")
	private static void copyCell(XSSFWorkbook book, XSSFCell oldCell, XSSFCell newCell,
			Map<Integer, XSSFCellStyle> styleMap) {
		if(styleMap != null) {     
	      int stHashCode = oldCell.getCellStyle().hashCode();     
	      XSSFCellStyle newCellStyle = styleMap.get(stHashCode);     
	      if(newCellStyle == null){     
	        newCellStyle = book.createCellStyle();     
	        newCellStyle.cloneStyleFrom(oldCell.getCellStyle());     
	        styleMap.put(stHashCode, newCellStyle);     
	      }     
	      newCell.setCellStyle(newCellStyle);   
	    }
		
		switch(oldCell.getCellType()) {     
		    case XSSFCell.CELL_TYPE_STRING:     
		      newCell.setCellValue(oldCell.getRichStringCellValue());     
		      break;     
		    case XSSFCell.CELL_TYPE_NUMERIC:     
		      newCell.setCellValue(oldCell.getNumericCellValue());     
		      break;     
		    case XSSFCell.CELL_TYPE_BLANK:     
		      newCell.setCellType(XSSFCell.CELL_TYPE_BLANK);     
		      break;     
		    case XSSFCell.CELL_TYPE_BOOLEAN:     
		      newCell.setCellValue(oldCell.getBooleanCellValue());     
		      break;     
		    case XSSFCell.CELL_TYPE_ERROR:     
		      newCell.setCellErrorValue(oldCell.getErrorCellValue());     
		      break;     
		    case XSSFCell.CELL_TYPE_FORMULA:     
		      newCell.setCellFormula(oldCell.getCellFormula());     
		      break;     
		    default:     
		      break;     
		  }
		
	}

}
