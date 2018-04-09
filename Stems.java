import java.util.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;

public class Stems {
	public static void main(String[] args) throws FileNotFoundException, IOException{
		Scanner reader = new Scanner(System.in);
		//Ask for file name
		System.out.println("What is the name of the file with the stored positions?");
		String file = reader.next();
		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));
		
		//Ask where the positions are stored
		System.out.println("On which sheet are the positions stored?");
		int sheetNum = reader.nextInt();
		XSSFSheet sheet = workbook.getSheetAt(sheetNum - 1);
		
		//Store the beginning and ending positions of each coding region
		double begin, end;
		double[][] codingRegions = new double[2][];
		System.out.println("How many coding regions are there?");
		int numCodes = reader.nextInt();
		codingRegions[0] = new double[numCodes];			//will store the start position of a coding region
		codingRegions[1] = new double[numCodes];			//will store the end position of a coding region
		for(int x = 1; x != numCodes + 1; x++){
			System.out.println("Where does coding region " + x + " begin?");
			begin = reader.nextInt();
			codingRegions[0][x-1] = begin;
			System.out.println("Where does coding region " + x + " end?");
			end = reader.nextInt();
			codingRegions[1][x-1] = end;
		}
		
		XSSFSheet results;
		XSSFRow row;
		XSSFCell cell;
		
		//Set up the sheet to store the results into
		results = workbook.createSheet("Stems");
		double sheetIndex = workbook.getSheetIndex("Stems");
		row = results.createRow(0);
		cell = row.createCell(0);
		cell.setCellValue("3-3 Stems");
		createCells(0,5,row);
		cell = row.createCell(6);
		cell.setCellValue("3-2 Stems");
		createCells(6,5,row);
		cell = row.createCell(12);
		cell.setCellValue("3-1 Stems");
		createCells(12,3,row);
		
		//Keep track of the row number for each column/stem type
		int rowCount33 = 1;
		int rowCount32 = 1;
		int rowCount31 = 1;
		int rowMax = 0;
		
		XSSFRow readRow = sheet.getRow(0);
		double num1, num2;
		double prev1 = readRow.getCell(0).getNumericCellValue();
		double prev2 = readRow.getCell(1).getNumericCellValue();
		double startHold1 = 0;						//Holds the starting position of side A of stem (top left of U)
		double startHold2 = 0;						//Holds the starting position of side B of stem (bottom right of U)
		int x = 0;
		boolean broken = false;
		
		for(readRow = sheet.getRow(x); readRow != null; readRow = sheet.getRow(x++)){
			num1 = readRow.getCell(0).getNumericCellValue(); 				//read the position number on one side of the stem
			num2 = readRow.getCell(1).getNumericCellValue(); 				//read the position number on the other side of the stem
			//check if num1 and num2 fall in any of the coding regions
			for(int y = 0; y < codingRegions[0].length; y++){
				if((num1 >= codingRegions[0][y]) && (num1 <= codingRegions[1][y])){
					if((num2 >= codingRegions[0][y] && (num2 <= codingRegions[1][y]))){
						if(startHold1 == 0){
							startHold1 = num1;
							startHold2 = num2;
						}
						if(Math.abs(num1-prev1) > 1 || Math.abs(num2-prev2) > 1){
							//Check if num1 and num2 codon pairing is 3-1, respectively
							for(int h = 0; h < 3 && (h + startHold1 <= prev1 && startHold2 - h >= prev2); h++){
								if(((((startHold1 + h) - codingRegions[0][y])% 3) + 1) == 3){
									if((((startHold2 - h) - codingRegions[0][y]) % 3) + 1 == 1){
										rowMax = findMax(rowCount33, rowCount32, rowCount31);
										if(rowCount31 >= rowMax){
											row = workbook.getSheetAt((int) sheetIndex).createRow(rowCount31+1);
										}else{
											row = workbook.getSheetAt((int) sheetIndex).getRow(rowCount31+1);
										}
										//Write the starting and ending positions for a 3-1 stem in the appropriate positions
										cell = row.createCell(12);
										cell.setCellValue(startHold1);
										cell = row.createCell(13);
										cell.setCellValue(prev1);
										cell = row.createCell(14);
										cell.setCellValue(prev2);
										cell = row.createCell(15);
										cell.setCellValue(startHold2);
										rowCount31++;
										startHold1 = num1;
										startHold2 = num2;
										broken = true;
										break;
									}
									else if((((startHold2 - h) - codingRegions[0][y]) % 3) + 1 == 2){
										rowMax = findMax(rowCount33, rowCount32, rowCount31);
										if(rowCount32 >= rowMax){
											row = workbook.getSheetAt((int) sheetIndex).createRow(rowCount32+1);
										}else{
											row = workbook.getSheetAt((int) sheetIndex).getRow(rowCount32+1);
										}
										//Write the starting and ending positions for a 3-2 stem in the appropriate positions
										cell = row.createCell(6);
										cell.setCellValue(startHold1);
										cell = row.createCell(7);
										cell.setCellValue(prev1);
										cell = row.createCell(8);
										cell.setCellValue(prev2);
										cell = row.createCell(9);
										cell.setCellValue(startHold2);
										rowCount32++;
										startHold1 = num1;
										startHold2 = num2;
										broken = true;
										break;
									}else{
										rowMax = findMax(rowCount33, rowCount32, rowCount31);
										if(rowCount33 >= rowMax){
											row = workbook.getSheetAt((int) sheetIndex).createRow(rowCount33+1);
										}else{
											row = workbook.getSheetAt((int) sheetIndex).getRow(rowCount33+1);
										}
										//Write the starting and ending positions for a 3-3 stem in the appropriate positions
										cell = row.createCell(0);
										cell.setCellValue(startHold1);
										cell = row.createCell(1);
										cell.setCellValue(prev1);
										cell = row.createCell(2);
										cell.setCellValue(prev2);
										cell = row.createCell(3);
										cell.setCellValue(startHold2);
										rowCount33++;
										startHold1 = num1;
										startHold2 = num2;
										broken = true;
										break;
									}
								}else if(((((startHold1 + h) - codingRegions[0][y])% 3) + 1) == 2){
									//Check if num1 and num2 pairing is 2-3, respectively
									if((((startHold2 - h) - codingRegions[0][y]) % 3) + 1 == 3){
										rowMax = findMax(rowCount33, rowCount32, rowCount31);
										if(rowCount32 >= rowMax){
											row = workbook.getSheetAt((int) sheetIndex).createRow(rowCount32+1);
										}else{
											row = workbook.getSheetAt((int) sheetIndex).getRow(rowCount32+1);
										}
										//Write the starting and ending positions for a 3-2 stem in the appropriate positions
										cell = row.createCell(6);
										cell.setCellValue(startHold1);
										cell = row.createCell(7);
										cell.setCellValue(prev1);
										cell = row.createCell(8);
										cell.setCellValue(prev2);
										cell = row.createCell(9);
										cell.setCellValue(startHold2);
										rowCount32++;
										startHold1 = num1;
										startHold2 = num2;
										broken = true;
										break;
									}
								}else{
									//Check if num1 and num2 pairing is 1-3, respectively
									if((((startHold2 - h) - codingRegions[0][y]) % 3) + 1 == 3){
										rowMax = findMax(rowCount33, rowCount32, rowCount31);
										if(rowCount31 >= rowMax){
											row = workbook.getSheetAt((int) sheetIndex).createRow(rowCount31+1);
										}else{
											row = workbook.getSheetAt((int) sheetIndex).getRow(rowCount31+1);
										}
										//Write the starting and ending positions for a 3-1 stem in the appropriate positions
										cell = row.createCell(12);
										cell.setCellValue(startHold1);
										cell = row.createCell(13);
										cell.setCellValue(prev1);
										cell = row.createCell(14);
										cell.setCellValue(prev2);
										cell = row.createCell(15);
										cell.setCellValue(startHold2);
										rowCount31++;
										startHold1 = num1;
										startHold2 = num2;
										broken = true;
										break;
									}
								}
							}
							if(!broken){
								startHold1 = num1;
								startHold2 = num2;
							}
							broken = false;
						}
						//break;
					}
				}
			}
			prev1 = num1;
			prev2 = num2;
		}
		
		//Create the workbook
		workbook.write(new FileOutputStream(file));
		workbook.close();
		reader.close();
		
	}
	
	//Create empty cells
	public static void createCells(int cellNum, int emptyWanted, XSSFRow row){
		XSSFCell cell;
		for(int x = 1; x <= emptyWanted; x++){
			cell = row.createCell(cellNum+x);
		}
	}
	
	public static int findMax(int a, int b, int c){
		int max;
		if(a>=b){
			max = a;
		}else{
			max = b;
		}
		if(c>max){
			max = c;
		}
		return max;
	}
}

