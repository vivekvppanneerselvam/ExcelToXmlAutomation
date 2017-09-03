package com.exceltoxml.converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
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
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class BYOExtractor {
	private static final String OUTPUT_FILE = "output/output.txt";
	private static final String INPUT_FILE = "input/ file name should be here and placed in project input path";
	static Workbook  wb = null  ;
	static Sheet requestedSheet = null;
	static List<StringBuilder> xmlData  = new ArrayList<StringBuilder>();
	static DataFormatter formatter = new DataFormatter();
	static String definedEntryFileVal = INPUT_FILE.replaceAll("[input/abcdmodel.xlsx-]" , "");
	private static StringBuilder constXmlKeyValues = new StringBuilder();
	static int count = 0;
	static int appendCount = 0;
	private static StringBuilder tempBuilder = new StringBuilder();
	private static StringBuilder finalBuilder = new StringBuilder();
	//default
	static List<String> defaultValues = new ArrayList<String>();
	//exterior
	static boolean defaultFlg = false;
	static boolean paintsFlg = false;
	static boolean wheelsFlg = false;
	static boolean optionsPkgFlg = false;
	static boolean optionFactoryFlg = false;
	static boolean accessoriesFlg = false;
	static Map<Integer,String> KeyValueCol = new  HashMap<Integer,String>();
	static HashMap<Integer, String> optionsKeyValueCol = new  HashMap<Integer,String>();
	static boolean[] exteriorTitleFlg = {defaultFlg, paintsFlg, wheelsFlg, optionsPkgFlg, optionFactoryFlg, accessoriesFlg};
	static int defaults,paints,wheels,optionsPkg,optionFactory,accessories =0;
	static int[] exteriorRows = {defaults,paints,wheels,optionsPkg,optionFactory,accessories};
	static List<String> listOfPkgs = new ArrayList <String>();

	//interior
	static boolean colorFlg = false;
	static boolean trimflg = false;
	static boolean intOptnsPkgFlg = false;
	static boolean intOptnsFactoryFlg = false;
	static boolean intAccessoriesFlg = false;
	static boolean[] interiorTitleFlg = {colorFlg, trimflg, intOptnsPkgFlg, intOptnsFactoryFlg, intAccessoriesFlg};
	//entertainment
	static boolean entkeyStdFeaturesFlg = false;
	static boolean entOptnsPkgFlg = false;
	static boolean entOptnsFactoryFlg = false;
	static boolean entAccessoriesFlg = false;
	static boolean[] entertainmentTitleFlg = {entkeyStdFeaturesFlg, entOptnsPkgFlg, entOptnsFactoryFlg, entAccessoriesFlg};
	//performance
	static boolean perfkeyStdFeaturesFlg = false;
	static boolean perfOptnsPkgFlg = false;
	static boolean perfOptnsFactoryFlg = false;
	static boolean perfAccessoriesFlg = false;
	static boolean[] performanceTitleFlg = {perfkeyStdFeaturesFlg, perfOptnsPkgFlg, perfOptnsFactoryFlg, perfAccessoriesFlg};
	//service
	static boolean[] serviceTitleFlg = {};

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


	public static int[] exteriorRowAnalysis(Sheet reqSheet, int rowItr){

		for(int row =0; row<rowItr; row++){
			if(!isRowEmpty(reqSheet.getRow(row))){
				String titleName = reqSheet.getRow(row).getCell(0).getStringCellValue();
				if(titleName.equalsIgnoreCase("defaults")){
					exteriorRows[0] = row;
				}else if(titleName.equalsIgnoreCase("paints")){
					exteriorRows[1] = row;
				}else if(titleName.equalsIgnoreCase("wheels")){
					exteriorRows[2] = row;
				}else if(titleName.equalsIgnoreCase("options (packages)")){
					exteriorRows[3] = row;
				}else if(titleName.equalsIgnoreCase("options (factory installed)")){
					exteriorRows[4] = row;
				}else if(titleName.equalsIgnoreCase("accessories")){
					exteriorRows[5] = row;
				}
			}else{
				continue;
			}
		}
		return exteriorRows;
	}

	public static boolean[] exteriorTitleAnalysis(Sheet reqSheet, int row){
		String titleName = reqSheet.getRow(row).getCell(0).getStringCellValue();
		if(titleName.equalsIgnoreCase("defaults")){
			Arrays.fill(exteriorTitleFlg, false);
			exteriorTitleFlg[0] = true;
		}else if(titleName.equalsIgnoreCase("paints")){
			Arrays.fill(exteriorTitleFlg, false);
			exteriorTitleFlg[1] = true;
			count=0;
			appendCount =0;
		}else if(titleName.equalsIgnoreCase("wheels")){
			Arrays.fill(exteriorTitleFlg, false);
			exteriorTitleFlg[2] = true;
			count=0;
			appendCount =0;
		}else if(titleName.equalsIgnoreCase("options (packages)")){
			Arrays.fill(exteriorTitleFlg, false);
			exteriorTitleFlg[3] = true;
			count=0;
			appendCount =0;
		}else if(titleName.equalsIgnoreCase("options (factory installed)")){
			Arrays.fill(exteriorTitleFlg, false);
			exteriorTitleFlg[4] = true;
			count=0;
			appendCount =0;
		}else if(titleName.equalsIgnoreCase("accessories")){
			Arrays.fill(exteriorTitleFlg, false);
			exteriorTitleFlg[5] = true;
			count=0;
			appendCount =0;
		}
		return exteriorTitleFlg;
	}


	public static int switchStatementExec(int rowItr ,int row, String type, int typColInt , Sheet requestedSheet, String fieldTyp){
		int [] arrInt = new int[]{0,2,typColInt};
		int k;
		count =0;
		for(k=row;k<rowItr; k++){
			count++;
			tempBuilder = new StringBuilder();
			finalBuilder = new StringBuilder();
			if(!isRowEmpty(requestedSheet.getRow(k))){
				for (int intValue : arrInt){
					switch(intValue){
					case 0: tempBuilder.append(requestedSheet.getRow(k).getCell(0).getStringCellValue()); break;
					case 2: finalBuilder.append("<entry key="+"\"img"+type+"Resp"+fieldTyp+"1"+"\""+">"+ requestedSheet.getRow(k).getCell(2).getNumericCellValue()+"</entry>"); break;
					case 3: tempBuilder.append(requestedSheet.getRow(k).getCell(3).getStringCellValue()); break;
					case 4: tempBuilder.append(requestedSheet.getRow(k).getCell(4).getStringCellValue()); break;
					case 5: tempBuilder.append(requestedSheet.getRow(k).getCell(5).getStringCellValue()); break;
					case 6: tempBuilder.append(requestedSheet.getRow(k).getCell(6).getStringCellValue()); break;
					default:
					}
				}
				constXmlKeyValues.append("<entry key="+"\"price"+type+"Resp"+fieldTyp+(count)+"\""+">"+tempBuilder +"</entry>");
				constXmlKeyValues.append(finalBuilder);
			}else{
				continue;
			}
		}
		return k;
	}

	public static List<Integer> iteratePkgAnalysis(Sheet reqSheet, int rowItr, int lastRowOptionPkg, List<String> listOfPkgs){
		List<Integer> pkgRows = new ArrayList<Integer>();
		for(String pkg: listOfPkgs){
			System.out.println("tested:"+pkg);
		}

		for(String pkg :listOfPkgs){
			for(int row = lastRowOptionPkg; row<=rowItr; row++){
				if(!isRowEmpty(reqSheet.getRow(row))){
					String titleName = reqSheet.getRow(row).getCell(0).getStringCellValue();
					System.out.println(titleName);
					if(pkg.toLowerCase().replaceAll("package", "").trim().contains(titleName.toLowerCase().replaceAll("package", "").trim())){
						pkgRows.add(row);
						System.out.println("added"+row);
					}else if(titleName.contains("options (factory installed)")){
						pkgRows.add(row);
						System.out.println("added"+row);
					}

				}else{
					continue;
				}
			}
		}
		return pkgRows;
	}

	public static boolean checkDuplicate(List<String> pkgs, String value){
		boolean duplFlg = true;
		for(String pkg :pkgs){
			System.out.println(pkg);
			if(pkg.toLowerCase().replaceAll("package", "").trim().contains(value.toLowerCase().replaceAll("package", "").trim())){
				System.out.println("Duplicate");
				duplFlg = false;
				exteriorTitleFlg[3] = false;
				break;
			}
		}
		return duplFlg;
	}

	public static int returnListOfPkgs(int rowItr ,int row, Sheet requestedSheet){
		listOfPkgs = new ArrayList <String>();
		int rowVal = 0;
		for(int k=row; k<rowItr; k++){
			if(!isRowEmpty(requestedSheet.getRow(k))){
				if(!requestedSheet.getRow(k).getCell(0).getStringCellValue().equalsIgnoreCase("name")){
					if(listOfPkgs.isEmpty()){
						System.out.println("is Empty");
						listOfPkgs.add(requestedSheet.getRow(k).getCell(0).getStringCellValue());
					}else{
						System.out.println("is not Empty");
						System.out.println(checkDuplicate(listOfPkgs, requestedSheet.getRow(k).getCell(0).getStringCellValue()));
						if(checkDuplicate(listOfPkgs, requestedSheet.getRow(k).getCell(0).getStringCellValue())){
							listOfPkgs.add(requestedSheet.getRow(k).getCell(0).getStringCellValue());
						}else{
							rowVal = k;
							break;
						}
					}
				}else{
					continue;
				}
			}else{
				continue;
			}
		}
		return rowVal;
	}

	public static void switchOptionPackagesExec(int rowItr ,int row, HashMap<Integer, String> keyNdValue, Sheet requestedSheet, String sheetTyp){
		int [] rdColumn = new int []{};
		count= 0;
		int lastRowOptn = returnListOfPkgs(rowItr , row, requestedSheet);
		List<Integer>listPkgRows = iteratePkgAnalysis(requestedSheet,rowItr,lastRowOptn, listOfPkgs);
		for (int m : keyNdValue.keySet()){
			rdColumn = new int []{0, m};
			for(int k=row; k<rowItr; k++){
				count++;
				tempBuilder = new StringBuilder();
				finalBuilder = new StringBuilder();
				if(!isRowEmpty(requestedSheet.getRow(k))){
					if(requestedSheet.getRow(k).getCell(0).getStringCellValue().equalsIgnoreCase("name")){
						if(checkDuplicate(listOfPkgs, requestedSheet.getRow(k).getCell(0).getStringCellValue())){
							for(int intValue : rdColumn){
								switch(intValue){
								case 0:
									tempBuilder.append(requestedSheet.getRow(k).getCell(0).getStringCellValue());
									break;
								case 1:
									tempBuilder.append(requestedSheet.getRow(k).getCell(1).getStringCellValue());
									System.out.println(requestedSheet.getRow(k).getCell(1).getStringCellValue());
									if(requestedSheet.getRow(k).getCell(1).getStringCellValue()!="NA"){
										buildOptionPkgVals(count, listPkgRows.get(1), listPkgRows.get(1-1), keyNdValue.get(m), requestedSheet, sheetTyp);
										System.out.println("case 1");
									}
									break;
								case 2:
									tempBuilder.append(requestedSheet.getRow(k).getCell(2).getStringCellValue());
									System.out.println(requestedSheet.getRow(k).getCell(2).getStringCellValue());
									if(requestedSheet.getRow(k).getCell(2).getStringCellValue()!="NA"){
										System.out.println("case 2");   
										buildOptionPkgVals(count,listPkgRows.get(1), listPkgRows.get(1-1), keyNdValue.get(m), requestedSheet, sheetTyp);

									}break;
								case 3:
									tempBuilder.append(requestedSheet.getRow(k).getCell(3).getStringCellValue());
									if(requestedSheet.getRow(k).getCell(3).getStringCellValue()!="NA"){
										buildOptionPkgVals(count,listPkgRows.get(3), listPkgRows.get(3-1), keyNdValue.get(m), requestedSheet, sheetTyp);
									}break;
								case 4:
									tempBuilder.append(requestedSheet.getRow(k).getCell(4).getStringCellValue());
									if(requestedSheet.getRow(k).getCell(4).getStringCellValue()!="NA"){
										buildOptionPkgVals(count,listPkgRows.get(4), listPkgRows.get(4-1), keyNdValue.get(m), requestedSheet, sheetTyp);
									}break;
								default:
								}
							}
						}else{
							continue;
						}
					}else{
						break;
					}
				} else{
					continue;
				}
			}
		}

	}

	public static void buildOptionPkgVals(int count, int endpt, int strtPt, String type, Sheet reqSheet, String sheetTyp){
		for(int k=strtPt; k<endpt; k++){
			finalBuilder = new StringBuilder();
			if(!isRowEmpty(reqSheet.getRow(k))){
				int colItr = reqSheet.getRow(k).getLastCellNum();
				for(int l=0; l<=colItr; l++){
					switch(l){
					case 0: constXmlKeyValues.append("<entry key="+"\"txt"+type+"Resp_"+sheetTyp+"PackageFeatures"+(count)+"_1"+"\""+">"+ reqSheet.getRow(k).getCell(l).getStringCellValue()+"</entry>"); break;
					case 1: constXmlKeyValues.append("<entry key="+"\"link"+type+"Resp_"+sheetTyp+"PackageFeatures"+(count)+"_1"+"\""+">"+ reqSheet.getRow(k).getCell(l).getStringCellValue()+"</entry>"); break;
					case 2: constXmlKeyValues.append("<entry key="+"\"desc"+type+"Resp_"+sheetTyp+"PackageFeatures"+(count)+"_1"+"\""+">"+ reqSheet.getRow(k).getCell(l).getStringCellValue()+"</entry>"); break;
					case 3: constXmlKeyValues.append("<entry key="+"\"disclaimer"+type+"Resp_"+sheetTyp+"PackageFeatures"+(count)+"_1"+"\""+">"+ reqSheet.getRow(k).getCell(l).getStringCellValue()+"</entry>"); break;
					case 4: constXmlKeyValues.append("<entry key="+"\"img"+type+"Resp_"+sheetTyp+"PackageFeatures"+(count)+"_1"+"\""+">//img[contains(@src,'"+ reqSheet.getRow(k).getCell(l).getStringCellValue().replaceAll("[.jpg]", "")+"')]</entry>");
					constXmlKeyValues.append("<entry key="+"\"ovimg"+type+"Resp_"+sheetTyp+"PackageFeatures"+(count)+"_1"+"\""+">//img[contains(@src,'"+ reqSheet.getRow(k).getCell(l).getStringCellValue().replaceAll("[.jpg]", "")+"')]</entry>");
					break;
					default:
					}
				}
			}else{
				continue;
			}
		}

	}

	public static int standaloneOptionSwitchExec(int row, Sheet reqSheet, String sheetTyp){
		int rowItr = reqSheet.getPhysicalNumberOfRows();
		int rowCount =0;
		for(String type: defaultValues){
			for(int j = row; j<rowItr; j++){
				if(!isRowEmpty(reqSheet.getRow(j))){
					exteriorTitleFlg = exteriorTitleAnalysis(reqSheet,j);
					int colItr = reqSheet.getRow(j).getPhysicalNumberOfCells();
					if(exteriorTitleFlg[4]){
						for(int k=0;k<colItr; k++){
							if(reqSheet.getRow(j).getCell(0).getStringCellValue()!= null && reqSheet.getRow(j).getCell(0).getStringCellValue().trim()!="" &&
									reqSheet.getRow(j).getCell(0).getCellType() != Cell.CELL_TYPE_BLANK){
								if(reqSheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase("options (factory installed)")){
									continue;
								}else{
									if(reqSheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase("name")){
										break;
									}else{
										switch(k){
										case 0: constXmlKeyValues.append("<entry key="+"\"link"+type+"_"+sheetTyp+"Options1"+"\""+">link= "+ reqSheet.getRow(j).getCell(k).getStringCellValue()+"</entry> "); break;
										case 1: constXmlKeyValues.append("<entry key="+"\"txt"+type+"_"+sheetTyp+"Options1"+"\""+">"+reqSheet.getRow(j).getCell(0).getStringCellValue()+ reqSheet.getRow(j).getCell(k).getStringCellValue()+"</entry>"); break;
										case 2: constXmlKeyValues.append("<entry key="+"\"desc"+type+"_"+sheetTyp+"Options1"+"\""+">"+ reqSheet.getRow(j).getCell(k).getStringCellValue()+"</entry>"); break;
										case 3: constXmlKeyValues.append("<entry key="+"\"disclaimer"+type+"_"+sheetTyp+"Options1"+"\""+">"+ reqSheet.getRow(j).getCell(k).getStringCellValue()+"</entry>"); break;
										case 4: constXmlKeyValues.append("<entry key="+"\"ovimg"+type+"_"+sheetTyp+"Options1"+"\""+">//img[contains(@src,'"+ reqSheet.getRow(j).getCell(k).getStringCellValue().replaceAll("[.jpg]", "")+"')]</entry>");
										constXmlKeyValues.append("<entry key="+"\"img"+type+"_"+sheetTyp+"Options1"+"\""+">//img[contains(@src,'"+ reqSheet.getRow(j).getCell(k).getStringCellValue().replaceAll("[.jpg]", "")+"')]</entry>");
										break;
										default:
										}
									}
								}
							}
						}
					}else{
						break;
					}
				}else{
					continue;
				}
			}
		}
		return rowCount;
	}

	public static int accessoriesSwitchExec(int row, Sheet reqSheet, String sheetTyp ){
		int rowItr = reqSheet.getPhysicalNumberOfRows();
		int rowCount =0;
		for(String type: defaultValues){
			for(int j = row; j<rowItr; j++){
				if(!isRowEmpty(reqSheet.getRow(j))){
					exteriorTitleFlg = exteriorTitleAnalysis(reqSheet,j);
					int colItr = reqSheet.getRow(j).getPhysicalNumberOfCells();
					if(exteriorTitleFlg[5]){
						for(int k=0;k<colItr; k++){
							if(reqSheet.getRow(j).getCell(0).getStringCellValue()!= null && reqSheet.getRow(j).getCell(0).getStringCellValue().trim()!="" &&
									reqSheet.getRow(j).getCell(0).getCellType() != Cell.CELL_TYPE_BLANK){
								if(reqSheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase("options (factory installed)")){
									continue;
								}else{
									if(reqSheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase("name")){
										break;
									}else{
										switch(k){
										case 0: constXmlKeyValues.append("<entry key="+"\"linkResp"+type+"_"+sheetTyp+"Accessories1"+"\""+">link= "+ reqSheet.getRow(j).getCell(k).getStringCellValue()+"</entry> "); break;
										case 1: constXmlKeyValues.append("<entry key="+"\"txtResp"+type+"_"+sheetTyp+"Accessories1"+"\""+">"+reqSheet.getRow(j).getCell(0).getStringCellValue()+ reqSheet.getRow(j).getCell(k).getStringCellValue()+"</entry>"); break;
										case 2: constXmlKeyValues.append("<entry key="+"\"descResp"+type+"_"+sheetTyp+"Accessories1"+"\""+">"+ reqSheet.getRow(j).getCell(k).getStringCellValue()+"</entry>"); break;
										case 3: constXmlKeyValues.append("<entry key="+"\"disclaimerResp"+type+"_"+sheetTyp+"Accessories1"+"\""+">"+ reqSheet.getRow(j).getCell(k).getStringCellValue()+"</entry>"); break;
										case 4: constXmlKeyValues.append("<entry key="+"\"ovimgResp"+type+"_"+sheetTyp+"Accessories1"+"\""+">//img[contains(@src,'"+ reqSheet.getRow(j).getCell(k).getStringCellValue().replaceAll("[.jpg]", "")+"')]</entry>");
										constXmlKeyValues.append("<entry key="+"\"imgResp"+type+"_"+sheetTyp+"Accessories1"+"\""+">//img[contains(@src,'"+ reqSheet.getRow(j).getCell(k).getStringCellValue().replaceAll("[.jpg]", "")+"')]</entry>");
										break;
										default:
										}
									}
								}
							}
						}
					}else{
						break;
					}
				}else{
					continue;
				}
			}
		}
		return rowCount;
	}

	public static void main(String[] args) throws ClassNotFoundException {
		try {
			BufferedWriter txtFile = null;
			FileInputStream inputFile = new FileInputStream(new File(INPUT_FILE));
			wb = new XSSFWorkbook(inputFile);
			for(int i=0;i<=5;i++){
				requestedSheet = wb.getSheet(wb.getSheetName(i));
				if(wb.getSheetName(i).trim().equalsIgnoreCase("exterior")){
					int rowItr = requestedSheet.getPhysicalNumberOfRows();
					int colItr = requestedSheet.getRow(0).getLastCellNum();
					System.out.println("Exterior");
					exteriorRowAnalysis(requestedSheet, rowItr);
					for(int j=0; j<rowItr; j++){
						if(!isRowEmpty(requestedSheet.getRow(j))){
							exteriorTitleFlg = exteriorTitleAnalysis(requestedSheet,j);
							if(exteriorTitleFlg[0]){
								if(requestedSheet.getRow(j).getCell(0).getStringCellValue()!= null  && requestedSheet.getRow(j).getCell(0).getStringCellValue() !="" && requestedSheet.getRow(j).getCell(0).getCellType() != Cell.CELL_TYPE_BLANK){
									if(requestedSheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase("name") || requestedSheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase("defaults")){
										continue;
									}else{
										defaultValues.add(requestedSheet.getRow(j).getCell(0).getStringCellValue().replaceAll(" ",""));
									}
								}else{
									continue;
								}
							}else if(exteriorTitleFlg[1]){
								for(int k=0;k<colItr; k++){
									if(requestedSheet.getRow(j).getCell(0).getStringCellValue()!= null && requestedSheet.getRow(j).getCell(0).getStringCellValue().trim()!="" &&
											requestedSheet.getRow(j).getCell(0).getCellType() != Cell.CELL_TYPE_BLANK){
										if(requestedSheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase("paints")){
											continue;
										}else{
											if(requestedSheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase("name")){
												for(String value : defaultValues){
													if(requestedSheet.getRow(j).getCell(k).getStringCellValue().toLowerCase().trim().contains(value.toLowerCase().replaceAll("amg", ""))){
														System.out.println("entered:-"+value);
														KeyValueCol.put(k, requestedSheet.getRow(j).getCell(k).getStringCellValue().toUpperCase());
													}else{
														System.out.println("entered cell value:-"+requestedSheet.getRow(j).getCell(k).getStringCellValue());
													}
												}
											}else{
												int newRow = 0;

												for(int m : KeyValueCol.keySet()){
													System.out.println(KeyValueCol.get(m));
													newRow = switchStatementExec(exteriorRows[2],j, KeyValueCol.get(m), m, requestedSheet, "ExteriorColor");
													System.out.println(newRow);
												}
												exteriorTitleAnalysis(requestedSheet,newRow);
												j = newRow;
											}
										}
									}else{
										continue;
									}
								}
							}else if(exteriorTitleFlg[2]){
								for(int k=0;k<colItr; k++){
									if(requestedSheet.getRow(j).getCell(0).getStringCellValue()!= null && requestedSheet.getRow(j).getCell(0).getStringCellValue().trim()!="" &&
											requestedSheet.getRow(j).getCell(0).getCellType() != Cell.CELL_TYPE_BLANK){
										if(requestedSheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase("wheels")){
											continue;
										}else{
											if(requestedSheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase("name")){
												for(String value : defaultValues){
													if(requestedSheet.getRow(j).getCell(k).getStringCellValue().toLowerCase().trim().contains(value.toLowerCase().replaceAll("amg", ""))){
														System.out.println("entered:-"+value);
														KeyValueCol.put(k, requestedSheet.getRow(j).getCell(k).getStringCellValue().toUpperCase());
													}else{
														System.out.println("entered cell value:-"+requestedSheet.getRow(j).getCell(k).getStringCellValue());
													}
												}
											}else{
												int newRow = 0;
												for(int m : KeyValueCol.keySet()){
													System.out.println(KeyValueCol.get(m));
													newRow = switchStatementExec(exteriorRows[3],j, KeyValueCol.get(m), m, requestedSheet, "ExteriorWheel");
													System.out.println(newRow);
												}
												exteriorTitleAnalysis(requestedSheet,newRow);
												j = newRow;
											}
										}
									}else{
										continue;
									}
								}
							}else if(exteriorTitleFlg[3]){
								for(int k=0;k<colItr; k++){
									if(requestedSheet.getRow(j).getCell(0).getStringCellValue()!= null && requestedSheet.getRow(j).getCell(0).getStringCellValue().trim()!="" &&
											requestedSheet.getRow(j).getCell(0).getCellType() != Cell.CELL_TYPE_BLANK){
										if(requestedSheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase("options (packages)")){
											continue;
										}else{
											if(requestedSheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase("name")){
												for(String value : defaultValues){
													if(requestedSheet.getRow(j).getCell(k).getStringCellValue().toLowerCase().trim().replaceAll(" ", "").contains(value.toLowerCase().replaceAll("amg", ""))){
														optionsKeyValueCol.put(k, requestedSheet.getRow(j).getCell(k).getStringCellValue().toUpperCase());
													}
												}
											}else{
											}
										}
									}
								}
								switchOptionPackagesExec(exteriorRows[4],j, optionsKeyValueCol, requestedSheet, "Ext");
							}else if(exteriorTitleFlg[4]){
								standaloneOptionSwitchExec(j,requestedSheet,"Exterior");
							}else if(exteriorTitleFlg[5]){
								accessoriesSwitchExec(j,requestedSheet,"Ext");
							}
						}
					}
				}else if(wb.getSheetName(i).trim().equalsIgnoreCase("interior")){
					int rowItr = requestedSheet.getPhysicalNumberOfRows();
					int colItr = requestedSheet.getRow(0).getLastCellNum();
					System.out.println("Interior");

				}else if(wb.getSheetName(i).trim().equalsIgnoreCase("entertainment")){
					int rowItr = requestedSheet.getPhysicalNumberOfRows();
					int colItr = requestedSheet.getRow(0).getLastCellNum();
					System.out.println("Entertainment");
				}else if(wb.getSheetName(i).trim().equalsIgnoreCase("performance")){
					int rowItr = requestedSheet.getPhysicalNumberOfRows();
					int colItr = requestedSheet.getRow(0).getLastCellNum();
					System.out.println("Performance");
				}else if(wb.getSheetName(i).trim().equalsIgnoreCase("service")){
					int rowItr = requestedSheet.getPhysicalNumberOfRows();
					int colItr = requestedSheet.getRow(0).getLastCellNum();
					System.out.println("Service");
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