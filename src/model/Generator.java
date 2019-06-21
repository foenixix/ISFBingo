package model;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Generator {
	
	final static Random rand = new Random(69);
	
	
	private final static int MIN_NUMBER = 1;
	private final static int MAX_NUMBER = 90;
	private final static int SQUARE_LENGTH = 5;
	private final static int SQUARES = 102;
	
	public static void main(String[] args) throws IOException {
		ArrayList<ArrayList<Integer>> numberList = generateNumbers(MIN_NUMBER,MAX_NUMBER,SQUARE_LENGTH*SQUARE_LENGTH,SQUARES);
		System.out.println(numberList);
		generateCards(numberList);
	}

	private static void generateCards(ArrayList<ArrayList<Integer>> numberList) throws IOException {
		XSSFWorkbook book = new XSSFWorkbook();
		Styles styles = new Styles(book);
		XSSFSheet sheet = book.createSheet();
		
		
		for(int page=0; page<SQUARES/6; page++) {
			List<ArrayList<Integer>> sixSquares = numberList.subList(page*6, page*6+6);
			for(int pageRow = 0; pageRow<15; pageRow+=5) {
				writeTopLine(sheet, styles, page*17+pageRow, 0, sixSquares.get(pageRow/5));
				writeTopLine(sheet, styles, page*17+pageRow, 5, sixSquares.get(3+pageRow/5));
				writeMiddleLines(sheet, styles, page*17+pageRow+1, 0, sixSquares.get(pageRow/5));
				writeMiddleLines(sheet, styles, page*17+pageRow+1, 5, sixSquares.get(3+pageRow/5));
				writeBottomLine(sheet, styles, page*17+pageRow+4, 0, sixSquares.get(pageRow/5));
				writeBottomLine(sheet, styles, page*17+pageRow+4, 5, sixSquares.get(3+pageRow/5));
			}
		}
		
		FileOutputStream stream = new FileOutputStream("output.xlsx");
		book.write(stream);
		stream.close();
		book.close();
		System.out.println("done");
	}

	private static void writeTopLine(XSSFSheet sheet, Styles styles, int rowNum, int startColumn, ArrayList<Integer> square) {
		XSSFRow row = sheet.getRow(rowNum);
		if(row ==null) {
			row = sheet.createRow(rowNum);
		}
		
		row.createCell(startColumn).setCellValue(square.get(0));
		row.getCell(startColumn).setCellStyle(styles.topLeft);
		for(int i = 1; i < 4; i++) {
			row.createCell(startColumn+i).setCellValue(square.get(i));
			row.getCell(startColumn+i).setCellStyle(styles.top);
		}
		row.createCell(startColumn+4).setCellValue(square.get(4));
		row.getCell(startColumn+4).setCellStyle(styles.topRight);
		
	}

	private static void writeMiddleLines(XSSFSheet sheet, Styles styles, int startRow, int startColumn, ArrayList<Integer> square) {
		for(int i=0;i<3;i++) {
			XSSFRow row = sheet.getRow(startRow+i);
			if(row ==null) {
				row = sheet.createRow(startRow+i);
			}
			
			row.createCell(startColumn).setCellValue(square.get(5+i*5));
			row.getCell(startColumn).setCellStyle(styles.left);
			for(int j = 1; j < 4; j++) {
				row.createCell(startColumn+j).setCellValue(square.get(5+i*5+j));
				row.getCell(startColumn+j).setCellStyle(styles.middle);
			}
			row.createCell(startColumn+4).setCellValue(square.get(9+i*5));
			row.getCell(startColumn+4).setCellStyle(styles.right);

		}
	}
	

	private static void writeBottomLine(XSSFSheet sheet, Styles styles, int rowNum, int startColumn, ArrayList<Integer> square) {
		XSSFRow row = sheet.getRow(rowNum);
		if(row ==null) {
			row = sheet.createRow(rowNum);
		}
		
		row.createCell(startColumn).setCellValue(square.get(20));
		row.getCell(startColumn).setCellStyle(styles.bottomLeft);
		for(int i = 1; i < 4; i++) {
			row.createCell(startColumn+i).setCellValue(square.get(20+i));
			row.getCell(startColumn+i).setCellStyle(styles.bottom);
		}
		row.createCell(startColumn+4).setCellValue(square.get(24));
		row.getCell(startColumn+4).setCellStyle(styles.bottomRight);
	}

	private static ArrayList<ArrayList<Integer>> generateNumbers(int min, int max, int nPerSquare, int nOfSquares) {
		ArrayList<ArrayList<Integer>> result = new ArrayList<>();
		for(int i=0; i<nOfSquares; i++) {
			ArrayList<Integer> square = new ArrayList<>();
			for(int j=0; j<nPerSquare; j++) {
				int newNumber = getRandomBetween(min,max);
				while(square.contains(newNumber)) {
					newNumber = getRandomBetween(min, max);
				}
				square.add(newNumber);
			}
			result.add(square);
		}
		return result;
	}

	private static int getRandomBetween(int min, int max) {
		return rand.nextInt(max)+min;
	}
	
	private static class Styles{
		
		public final CellStyle topLeft;
		public final CellStyle topRight;
		public final CellStyle bottomLeft;
		public final CellStyle bottomRight;
		public final CellStyle top;
		public final CellStyle bottom;
		public final CellStyle left;
		public final CellStyle right;
		public final CellStyle middle;
		
		public Styles(XSSFWorkbook book) {
			Font font = book.createFont();
			font.setBold(true);
			font.setFontHeightInPoints((short) 35);
			
			topLeft = book.createCellStyle();
			topLeft.setFont(font);
			topLeft.setAlignment(CellStyle.ALIGN_CENTER);
			topLeft.setBorderTop(CellStyle.BORDER_THICK);
			topLeft.setBorderLeft(CellStyle.BORDER_THICK);
			topLeft.setBorderBottom(CellStyle.BORDER_THIN);
			topLeft.setBorderRight(CellStyle.BORDER_THIN);
			
			topRight = book.createCellStyle();
			topRight.setFont(font);
			topRight.setAlignment(CellStyle.ALIGN_CENTER);
			topRight.setBorderTop(CellStyle.BORDER_THICK);
			topRight.setBorderLeft(CellStyle.BORDER_THIN);
			topRight.setBorderBottom(CellStyle.BORDER_THIN);
			topRight.setBorderRight(CellStyle.BORDER_THICK);
			
			bottomLeft = book.createCellStyle();
			bottomLeft.setFont(font);
			bottomLeft.setAlignment(CellStyle.ALIGN_CENTER);
			bottomLeft.setBorderTop(CellStyle.BORDER_THIN);
			bottomLeft.setBorderLeft(CellStyle.BORDER_THICK);
			bottomLeft.setBorderBottom(CellStyle.BORDER_THICK);
			bottomLeft.setBorderRight(CellStyle.BORDER_THIN);

			bottomRight = book.createCellStyle();
			bottomRight.setFont(font);
			bottomRight.setAlignment(CellStyle.ALIGN_CENTER);
			bottomRight.setBorderTop(CellStyle.BORDER_THIN);
			bottomRight.setBorderLeft(CellStyle.BORDER_THIN);
			bottomRight.setBorderBottom(CellStyle.BORDER_THICK);
			bottomRight.setBorderRight(CellStyle.BORDER_THICK);
			
			top = book.createCellStyle();
			top.setFont(font);
			top.setAlignment(CellStyle.ALIGN_CENTER);
			top.setBorderTop(CellStyle.BORDER_THICK);
			top.setBorderLeft(CellStyle.BORDER_THIN);
			top.setBorderBottom(CellStyle.BORDER_THIN);
			top.setBorderRight(CellStyle.BORDER_THIN);

			bottom = book.createCellStyle();
			bottom.setFont(font);
			bottom.setAlignment(CellStyle.ALIGN_CENTER);
			bottom.setBorderTop(CellStyle.BORDER_THIN);
			bottom.setBorderLeft(CellStyle.BORDER_THIN);
			bottom.setBorderBottom(CellStyle.BORDER_THICK);
			bottom.setBorderRight(CellStyle.BORDER_THIN);

			left = book.createCellStyle();
			left.setFont(font);
			left.setAlignment(CellStyle.ALIGN_CENTER);
			left.setBorderTop(CellStyle.BORDER_THIN);
			left.setBorderLeft(CellStyle.BORDER_THICK);
			left.setBorderBottom(CellStyle.BORDER_THIN);
			left.setBorderRight(CellStyle.BORDER_THIN);

			right = book.createCellStyle();
			right.setFont(font);
			right.setAlignment(CellStyle.ALIGN_CENTER);
			right.setBorderTop(CellStyle.BORDER_THIN);
			right.setBorderLeft(CellStyle.BORDER_THIN);
			right.setBorderBottom(CellStyle.BORDER_THIN);
			right.setBorderRight(CellStyle.BORDER_THICK);

			middle = book.createCellStyle();
			middle.setFont(font);
			middle.setAlignment(CellStyle.ALIGN_CENTER);
			middle.setBorderTop(CellStyle.BORDER_THIN);
			middle.setBorderLeft(CellStyle.BORDER_THIN);
			middle.setBorderBottom(CellStyle.BORDER_THIN);
			middle.setBorderRight(CellStyle.BORDER_THIN);
		}
	}
	
}
