package com.exceltoxml.converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ModelConverter {

	private static final String OUTPUT_FILE = "output/ConvertedXML.txt";
	private static final String INPUT_FILE = "input/abcdMY15-E400C-model.xlsx";
	static Workbook  wb = null  ;
	static Sheet requestedSheet = null;
	static List<StringBuilder> xmlData  = new ArrayList<StringBuilder>();
	static String definedEntryFileVal = INPUT_FILE.replaceAll("[input/abcdmodel.xlsx-]" , "");
	//Features initialize
	static boolean performanceFlg = false;
	static boolean design = false;
	static boolean safety = false;
	static boolean comfortConvFlg = false;
	static boolean audioNdEntMnt = false;
	static boolean  booleanCheckFlg = false;
	static boolean accessories = false;
	//common 
	private static StringBuilder constXmlKeyValues = new StringBuilder();
	static String type = null;
	//Specifications initialize 
	static boolean dimensionsFlg = false;
	static boolean keyFeaturesFlg =false;
	static boolean moreExtDmnsFlg =false;
	static boolean intDmnsFlg =false;
	static boolean booleanValFlg =false;
	static boolean engineDrivePerfFlg =false;
	static boolean chasistransCtrlFlg = false;
	static boolean brakesWheelTireFlg = false;
	static boolean safetyNdSecurityFlg = false;
	private static StringBuilder tempBuilder = new StringBuilder();

	static boolean[] specificationsTitleFlg = new boolean[]{dimensionsFlg, keyFeaturesFlg, moreExtDmnsFlg, intDmnsFlg, 
			booleanValFlg, engineDrivePerfFlg, chasistransCtrlFlg, brakesWheelTireFlg, safetyNdSecurityFlg}; 

	public static boolean isRowEmpty(Row row) {
		for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
			Cell cell = row.getCell(c);
			if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK ){
				switch(cell.getCellType()){
				case Cell.CELL_TYPE_BOOLEAN:
					return false;
				case Cell.CELL_TYPE_STRING:
					if(!cell.getStringCellValue().trim().equalsIgnoreCase("") && cell.getStringCellValue() != null){
						return false;
					}else{
						return true;
					}
				case Cell.CELL_TYPE_NUMERIC:
					return false;
				}
			}  
		}
		return true;
	}

	public static void switchStatementExec(int col,int row, String type, Sheet requestedSheet ){
		for(int k=0;k<col; k++){
			if(requestedSheet.getRow(row).getCell(k).getStringCellValue()!= null && requestedSheet.getRow(row).getCell(k).getStringCellValue().trim()!="" && requestedSheet.getRow(row).getCell(k).getCellType() != Cell.CELL_TYPE_BLANK){
				System.out.println("row:"+row+"col:"+k+" "+requestedSheet.getRow(row).getCell(k).getStringCellValue());
				switch(k){
				case 0: constXmlKeyValues.append("<entry key="+"\"title"+type+""+definedEntryFileVal+"Resp1\""+">"+ requestedSheet.getRow(row).getCell(k).getStringCellValue()+"</entry>"); break;
				case 1:
					constXmlKeyValues.append("<entry key="+"\"link"+type+""+definedEntryFileVal+"Resp1\""+">"+requestedSheet.getRow(row).getCell(k).getStringCellValue() +"</entry>"); break;
				case 2:
					constXmlKeyValues.append("<entry key="+"\"img"+type+""+definedEntryFileVal+"Resp1\""+">"+ requestedSheet.getRow(row).getCell(k).getStringCellValue()+"</entry>"); break;
				case 3:
					constXmlKeyValues.append("<entry key="+"\"price"+type+""+definedEntryFileVal+"Resp1\""+">"+ requestedSheet.getRow(row).getCell(k).getStringCellValue()+"</entry>"); break;
				case 4:
					constXmlKeyValues.append("<entry key="+"\"text"+type+""+definedEntryFileVal+"Resp1\""+">"+requestedSheet.getRow(row).getCell(k).getStringCellValue() +"</entry>"); break;
				case 5:
					constXmlKeyValues.append("<entry key="+"\"pkg"+type+""+definedEntryFileVal+"Resp1\""+">"+ requestedSheet.getRow(row).getCell(k).getStringCellValue()+"</entry>"); break;
				case 6:
					constXmlKeyValues.append("<entry key="+"\"disclaimer"+type+""+definedEntryFileVal+"Resp1\""+">"+ requestedSheet.getRow(row).getCell(k).getStringCellValue()+"</entry>");break;
				default:
				}
			}
		}
	}

	public static boolean[] specificationTitleAnalysis(Sheet reqSheet, int row){

		String titleName = reqSheet.getRow(row).getCell(0).getStringCellValue();
		System.out.println(titleName);
		if(titleName.equalsIgnoreCase("Dimensions")){
			Arrays.fill(specificationsTitleFlg, false);
			specificationsTitleFlg[0] = true;
		}else if(titleName.equalsIgnoreCase("Key Features")){
			Arrays.fill(specificationsTitleFlg, false);
			specificationsTitleFlg[1] = true;
		}else if(titleName.equalsIgnoreCase("More Exterior Dimensions")){
			Arrays.fill(specificationsTitleFlg, false);
			specificationsTitleFlg[2] = true;
		}else if(titleName.equalsIgnoreCase("Interior Dimensions")){
			Arrays.fill(specificationsTitleFlg, false);
			specificationsTitleFlg[3] = true;
		}else if(titleName.equalsIgnoreCase("Engine, Drivetrain, And Performance")){
			Arrays.fill(specificationsTitleFlg, false);
			specificationsTitleFlg[5] = true;
		}else if(titleName.equalsIgnoreCase("Chassis And TractionControl/stability Systems")){
			Arrays.fill(specificationsTitleFlg, false);
			specificationsTitleFlg[6] = true;
		}else if(titleName.equalsIgnoreCase("Brakes, Wheels, And Tires")){
			Arrays.fill(specificationsTitleFlg, false);
			specificationsTitleFlg[7] = true;
		}else if(titleName.equalsIgnoreCase("Safety And Security Systems")){
			Arrays.fill(specificationsTitleFlg, false);
			specificationsTitleFlg[8] = true;
		}
		return specificationsTitleFlg;

	}

	public static void main(String[] args) throws ClassNotFoundException {

		try {
			BufferedWriter txtFile = null;
			FileInputStream inputFile = new FileInputStream(new File(INPUT_FILE));
			wb = new XSSFWorkbook(inputFile);
			for(int i=0;i<=1;i++){
				requestedSheet = wb.getSheet(wb.getSheetName(i));
				if(wb.getSheetName(i).trim().equalsIgnoreCase("hero")){
					//int rowItr = requestedSheet.getPhysicalNumberOfRows();
					//int colItr = requestedSheet.getRow(0).getPhysicalNumberOfCells();
					System.out.println("Hero");
				}else if(wb.getSheetName(i).trim().equalsIgnoreCase("specifications")){
					int rowItr = requestedSheet.getPhysicalNumberOfRows();
					int colItr = requestedSheet.getRow(0).getLastCellNum();
					System.out.println("Specifications");
					for(int j=0; j<rowItr; j++){
						if(!isRowEmpty(requestedSheet.getRow(j))){
							if(requestedSheet.getRow(j).getCell(0).getCellType() == Cell.CELL_TYPE_BOOLEAN){
								System.out.println("Boolean Cell");
								specificationsTitleFlg[4]= true;
							}else{
								specificationsTitleFlg = specificationTitleAnalysis(requestedSheet, j);
							}
							if(specificationsTitleFlg[0]){
								for(int k=0;k<colItr; k++){
									if(requestedSheet.getRow(j).getCell(k).getStringCellValue()!= null && requestedSheet.getRow(j).getCell(k).getStringCellValue().trim()!="" && requestedSheet.getRow(j).getCell(k).getCellType() != Cell.CELL_TYPE_BLANK){
										//System.out.println(requestedSheet.getRow(j).getCell(0).getStringCellValue());
										if(requestedSheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase("Dimensions")){
											tempBuilder.append("<entry key="+"\"title"+definedEntryFileVal+"Resp_Specifications1/"+">"+requestedSheet.getRow(j).getCell(k).getStringCellValue()+"</entry>");
											System.out.println(requestedSheet.getRow(j).getCell(k).getStringCellValue());
										}else{
											k++;
											tempBuilder.append("<entry key="+"\"img"+definedEntryFileVal+"Resp_Specifications1/"+">"+requestedSheet.getRow(j).getCell(k).getStringCellValue()+"</entry>");
											System.out.println(requestedSheet.getRow(j).getCell(k).getStringCellValue());
										}
									}else{
										continue;
									}
								}
								constXmlKeyValues.append(tempBuilder);
							}else if(specificationsTitleFlg[1]){
								
							}else if(specificationsTitleFlg[2]){

							}else if(specificationsTitleFlg[3]){

							}else if(specificationsTitleFlg[4]){

							}else if(specificationsTitleFlg[5]){

							}
						}else{
							continue;
						}
					}
				}else if(wb.getSheetName(i).trim().equalsIgnoreCase("features")){
					int rowItr = requestedSheet.getPhysicalNumberOfRows();
					int colItr = requestedSheet.getRow(0).getPhysicalNumberOfCells();
					for(int j=0; j<rowItr; j++){
						String scenerioName = "";
						Cell toCheckCell = requestedSheet.getRow(j).getCell(0);
						if(!isRowEmpty(requestedSheet.getRow(j))) {
							if(toCheckCell.getCellType() == Cell.CELL_TYPE_BOOLEAN){
								System.out.println("Boolean Cell:"+toCheckCell.getCellType()+""+Cell.CELL_TYPE_BOOLEAN+""+toCheckCell.getBooleanCellValue());
								booleanCheckFlg = true;
								continue;
							}else{
								scenerioName = requestedSheet.getRow(j).getCell(0).getStringCellValue();
								System.out.println(scenerioName);
								if(scenerioName.equalsIgnoreCase("performance & handling")){
									performanceFlg = true;
									System.out.println("rowValue:"+j+""+requestedSheet.getRow(j).getCell(0).getStringCellValue());
									continue;
								}else if(scenerioName.equalsIgnoreCase("design")){
									System.out.println("rowValue:"+j+""+requestedSheet.getRow(j).getCell(0).getStringCellValue());
									design = true;
									performanceFlg = false;
									continue;
								}else if(scenerioName.equalsIgnoreCase("safety")){
									System.out.println("rowValue:"+j+""+requestedSheet.getRow(j).getCell(0).getStringCellValue());
									safety = true;
									design = false;
									continue;
								}else if(scenerioName.equalsIgnoreCase("comfort & convenience")){
									comfortConvFlg = true;
									safety = false;
								}
								else if(scenerioName.equalsIgnoreCase("audio & entertainment")){
									System.out.println("rowValue:"+j+""+requestedSheet.getRow(j).getCell(0).getStringCellValue());
									audioNdEntMnt = true;
									comfortConvFlg = false;
									continue;
								}else if(scenerioName.equalsIgnoreCase("accessories")){
									System.out.println("rowValue:"+j+""+requestedSheet.getRow(j).getCell(0).getStringCellValue());
									accessories = true;
									audioNdEntMnt  = false;
									continue;
								}else{
									accessories = false;
								}
							}
							if(performanceFlg){
								type = "PerformanceHandling";
								switchStatementExec(colItr, j, type, requestedSheet);
							}else if(design){
								type  = "Design"; 
								switchStatementExec(colItr, j, type, requestedSheet);
							}else if(safety){
								type = "Safety";
								switchStatementExec(colItr, j, type, requestedSheet);
							}else if(comfortConvFlg){
								type = "ComfortConvenience";
								switchStatementExec(colItr, j, type, requestedSheet);
							}else if(audioNdEntMnt){
								type ="AudioEntertainment";
								switchStatementExec(colItr, j, type, requestedSheet);
							}else if(accessories){
								type ="Accessories";
								switchStatementExec(colItr, j, type, requestedSheet);
							}else{
								type = "Packages";
								switchStatementExec(colItr, j, type, requestedSheet);
							}
						}else{
							continue;
						}
					}
				}else{
					System.out.println("Site share");
				}
				xmlData.add(constXmlKeyValues);
			}
			txtFile = new BufferedWriter(new FileWriter(OUTPUT_FILE));
			txtFile.write(constXmlKeyValues.toString());
			txtFile.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch(NoClassDefFoundError e){
			e.printStackTrace();
			System.out.println(e);
		}catch(IllegalArgumentException e){
			e.printStackTrace();
			System.out.println(e);
		}catch(NullPointerException e){
			e.printStackTrace();
			System.out.println(e);
		}catch(IllegalStateException e){
			e.printStackTrace();
			System.out.println(e);
		}

		System.out.println("Done");
	}


}