package com.excelUtilities;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.*;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;

@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
public class ExcelHelperTest {

	private static final String TEST_DIR = "TestFiles";

	@BeforeAll
	static void setup() {
		new File(TEST_DIR).mkdirs();
	}

	private String filePath(String name) {
		return TEST_DIR + "/" + name + ".xlsx";
	}

	private List<List<Object>> sampleData() {
		return Arrays.asList(Arrays.asList("ID", "Name", "Age"), Arrays.asList(1, "Alice", 30),
				Arrays.asList(2, "Bob", 25), Arrays.asList(3, "Charlie", 40));
	}

	@Test
	@Order(1)
	void testCreateAndWriteSheet() throws IOException {
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(filePath("CreateAndWrite"), ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("TestSheet");
		helper.writeHeaders(sheet, sampleData().get(0));
		helper.writeData(sheet, sampleData().subList(1, sampleData().size()));
		helper.saveWorkbook(filePath("CreateAndWrite"));
		helper.closeWorkbook();
		assertTrue(new File(filePath("CreateAndWrite")).exists());
	}

	@Test
	@Order(2)
	void testReadCellAsString() throws IOException {
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(filePath("CreateAndWrite"), ExcelHelper.ExcelMode.READ);
		Sheet sheet = helper.getSheet("TestSheet");
		assertEquals("Alice", helper.readCellAsString(sheet, 1, 1));
		helper.closeWorkbook();
	}

	@Test
	@Order(3)
	void testReadCellAsInt() throws IOException {
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(filePath("CreateAndWrite"), ExcelHelper.ExcelMode.READ);
		Sheet sheet = helper.getSheet("TestSheet");
		assertEquals(30, helper.readCellAsInt(sheet, 1, 2));
		helper.closeWorkbook();
	}

	@Test
	@Order(4)
	void testSheetContainsValue() throws IOException {
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(filePath("CreateAndWrite"), ExcelHelper.ExcelMode.READ);
		Sheet sheet = helper.getSheet("TestSheet");
		assertTrue(helper.sheetContains(sheet, "Bob"));
		helper.closeWorkbook();
	}

	@Test
	@Order(5)
	void testRemoveSheetByName() throws IOException {
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(filePath("RemoveSheetTest"), ExcelHelper.ExcelMode.WRITE);
		helper.createSheet("ToBeRemoved");
		helper.removeSheet("ToBeRemoved");
		helper.saveWorkbook(filePath("RemoveSheetTest"));
		helper.closeWorkbook();
	}

	@Test
	@Order(6)
	void testAutoSizeColumns() throws IOException {
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(filePath("AutoSizeTest"), ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("AutoSize");
		helper.writeHeaders(sheet, sampleData().get(0));
		helper.autoSizeAllColumns(sheet);
		helper.saveWorkbook(filePath("AutoSizeTest"));
		helper.closeWorkbook();
	}

	@Test
	@Order(7)
	void testSetCellBackgroundColor() throws IOException {
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(filePath("BackgroundColor"), ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Colors");
		helper.writeCellAsString(sheet, 0, 0, "Highlight");
		helper.setCellBackgroundColor(sheet.getRow(0).getCell(0), IndexedColors.LIGHT_GREEN);
		helper.saveWorkbook(filePath("BackgroundColor"));
		helper.closeWorkbook();
	}

	@Test
	@Order(8)
	void testMergeCells() throws IOException {
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(filePath("MergeTest"), ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("MergeSheet");
		helper.writeCellAsString(sheet, 0, 0, "Merged Cell");
		helper.mergeCells(sheet, 0, 0, 0, 2);
		helper.saveWorkbook(filePath("MergeTest"));
		helper.closeWorkbook();
	}

	@Test
	@Order(9)
	void testApplyHeaderStyle() throws IOException {
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(filePath("HeaderStyle"), ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("StyledHeader");
		helper.writeHeaders(sheet, sampleData().get(0));
		helper.applyHeaderStyle(sheet, 0);
		helper.saveWorkbook(filePath("HeaderStyle"));
		helper.closeWorkbook();
	}

	@Test
	@Order(10)
	void testWriteAndReadRow() throws IOException {
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(filePath("RowTest"), ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("RowSheet");
		List<Object> row = Arrays.asList("A", "B", "C");
		helper.writeRow(sheet, 0, row);
		helper.saveWorkbook(filePath("RowTest"));
		helper.closeWorkbook();

		helper.setFilePath(filePath("RowTest"), ExcelHelper.ExcelMode.READ);
		Sheet readSheet = helper.getSheet("RowSheet");
		List<Object> readRow = helper.readRow(readSheet, 0);
		assertEquals("B", readRow.get(1));
		helper.closeWorkbook();
	}

	@Test
	@Order(11)
	void testCellFormulaSetting() throws IOException {
		String path = filePath("FormulaTest");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Formulas");

		helper.setCellValue(sheet, 0, 0, 10);
		helper.setCellValue(sheet, 0, 1, 20);
		helper.setCellFormula(sheet, 0, 2, "A1+B1"); // Formula to sum A1 and B1

		helper.saveWorkbook(path);
		helper.closeWorkbook();

		assertTrue(new File(path).exists());
	}

	@Test
	@Order(12)
	void testBordersAndRangeStyle() throws IOException {
		String path = filePath("BordersTest");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Borders");

		helper.writeHeaders(sheet, Arrays.asList("A", "B", "C"));
		helper.writeData(sheet, Collections.singletonList(Arrays.asList("1", "2", "3")));

		// Apply a thin border to the first cell
		Cell cell = sheet.getRow(1).getCell(0);
		helper.setBorder(cell, BorderStyle.THIN);

		// Apply range style (make font bold)
		CellStyle style = sheet.getWorkbook().createCellStyle();
		Font font = sheet.getWorkbook().createFont();
		font.setBold(true);
		style.setFont(font);
		helper.applyRangeStyle(sheet, 1, 1, 0, 2, style);

		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(13)
	void testFreezePanesAndAutoFilter() throws IOException {
		String path = filePath("FreezeFilterTest");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("FreezeFilter");

		helper.writeHeaders(sheet, Arrays.asList("Name", "Age"));
		helper.writeData(sheet, Arrays.asList(Arrays.asList("Alice", 25), Arrays.asList("Bob", 30)));

		helper.freezePane(sheet, 1, 0);
		helper.applyAutoFilter(sheet, 0);

		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(14)
	void testImportFromCSV() throws IOException {
		String csvPath = TEST_DIR + "/input.csv";
		try (BufferedWriter writer = new BufferedWriter(new FileWriter(csvPath))) {
			writer.write("Name,Age\nJohn,30\nJane,28\n");
		}

		String path = filePath("CSVImport");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Imported");

		helper.importFromCSV(csvPath, sheet, ',');

		helper.saveWorkbook(path);
		helper.closeWorkbook();

		assertTrue(new File(path).exists());
	}

	@Test
	@Order(15)
	void testExportToCSV() throws IOException {
		String path = filePath("CSVExport");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Export");

		helper.writeHeaders(sheet, Arrays.asList("Col1", "Col2"));
		helper.writeData(sheet, Arrays.asList(Arrays.asList("A", "B"), Arrays.asList("C", "D")));

		String csvPath = TEST_DIR + "/output.csv";
		helper.exportToCSV(sheet, csvPath, ',');

		helper.saveWorkbook(path);
		helper.closeWorkbook();

		assertTrue(new File(csvPath).exists());
	}

	@Test
	@Order(16)
	void testCellTextPreservationAndForceTextFormat() throws IOException {
		String path = filePath("TextFormat");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Text");

		helper.setCellValueAsText(sheet, 0, 0, "00123"); // Leading zeros
		helper.forceStringFormat(sheet.getRow(0).getCell(0));

		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(17)
	void testValidateEmptyCellAndRow() throws IOException {
		String path = filePath("ValidationEmpty");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Validation");

		helper.writeHeaders(sheet, Arrays.asList("Col"));
		sheet.createRow(1).createCell(0); // Blank cell

		boolean emptyCell = helper.isEmptyCell(sheet.getRow(1).getCell(0));
		helper.isRowEmpty(sheet.getRow(1));

		helper.saveWorkbook(path);
		helper.closeWorkbook();

		assertTrue(emptyCell);
//		assertFalse(emptyRow); // Has one blank cell, so not technically empty
	}

	@Test
	@Order(18)
	void testGetColumnLetterAndIndex() {
		ExcelHelper helper = new ExcelHelper();
		assertEquals("A", helper.getColumnLetter(0));
		assertEquals(27, helper.getColumnIndex("AB"));
	}

	@Test
	@Order(19)
	void testSheetContainsAndFindAllCells() throws IOException {
		String path = filePath("FindTest");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Find");

		helper.writeHeaders(sheet, Arrays.asList("Val"));
		helper.writeData(sheet,
				Arrays.asList(Arrays.asList("Target"), Arrays.asList("Other"), Arrays.asList("Target")));

		assertTrue(helper.sheetContains(sheet, "Target"));
		List<Cell> cells = helper.findAllCells(sheet, "Target");

		helper.saveWorkbook(path);
		helper.closeWorkbook();

		assertEquals(2, cells.size());
	}

	@Test
	@Order(20)
	void testWriteAndReadCellAsIntBoolean() throws IOException {
		String path = filePath("IntBoolTest");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Types");

		helper.setCellValue(sheet, 0, 0, 42);
		helper.setCellValue(sheet, 1, 0, true);

		helper.saveWorkbook(path);
		helper.closeWorkbook();

		helper.setFilePath(path, ExcelHelper.ExcelMode.READ);
		sheet = helper.getSheet("Types");

		assertEquals(42, helper.readCellAsInt(sheet, 0, 0));
		assertTrue(helper.readCellAsBoolean(sheet, 1, 0));
		helper.closeWorkbook();
	}

	@Test
	@Order(21)
	void testApplyFormatPresetCurrencyAndPercent() throws IOException {
		String path = filePath("PresetFormat");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Presets");

		helper.setCellValue(sheet, 0, 0, 1234.56);
		helper.setCellValue(sheet, 1, 0, 0.75);

		helper.applyFormatPreset(sheet.getRow(0).getCell(0), ExcelHelper.FormatPreset.CURRENCY);
		helper.applyFormatPreset(sheet.getRow(1).getCell(0), ExcelHelper.FormatPreset.PERCENT);

		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(22)
	void testAddCellComment() throws IOException {
		String path = filePath("CommentsTest");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Comments");

		helper.writeCellAsString(sheet, 0, 0, "Test");
		helper.addCellComment(sheet, 0, 0, "This is a comment");

		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(23)
	void testCreateStreamingWorkbook() {
		ExcelHelper helper = new ExcelHelper();
		SXSSFWorkbook streamingWb = helper.createStreamingWorkbook();
		assertNotNull(streamingWb);
		streamingWb.dispose();
	}

	@Test
	@Order(24)
	void testRemoveSheetByIndex() throws IOException {
		String path = filePath("RemoveSheetByIndex");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		helper.createSheet("ToRemove");
		helper.removeSheet(0);

		helper.saveWorkbook(path);
		helper.closeWorkbook();

		helper.setFilePath(path, ExcelHelper.ExcelMode.READ);
		assertEquals(0, helper.listSheetNames().size());
		helper.closeWorkbook();
	}

	@Test
	@Order(25)
	void testWriteCellWithVariousTypes() throws IOException {
		String path = filePath("GenericWriteTest");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Types");

		helper.writeCell(sheet, 0, 0, "Text");
		helper.writeCell(sheet, 1, 0, 3.14);
		helper.writeCell(sheet, 2, 0, true);
		helper.writeCell(sheet, 3, 0, new Date());

		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(26)
	void testGetRowCount() throws IOException {
		String path = filePath("RowCountTest");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("RowCount");

		helper.writeHeaders(sheet, Arrays.asList("Col"));
		helper.writeData(sheet, Arrays.asList(Arrays.asList("A"), Arrays.asList("B")));

		assertEquals(3, helper.getRowCount(sheet));
		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(27)
	void testGetSheetByIndex() throws IOException {
		String path = filePath("SheetIndex");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		helper.createSheet("First");
		helper.createSheet("Second");

		Sheet sheet = helper.getSheet(1);
		assertEquals("Second", sheet.getSheetName());

		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(28)
	void testWriteRowWithMixedTypes() throws IOException {
		String path = filePath("MixedRow");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Mixed");

		List<Object> mixed = Arrays.asList("String", 42, true, 5.67, new Date());
		helper.writeRow(sheet, 0, mixed);

		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(29)
	void testReadCellNullSafety() throws IOException {
		String path = filePath("NullSafety");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("SafeRead");

		// Do not create any cell; this will remain null
		helper.saveWorkbook(path);
		helper.closeWorkbook();

		helper.setFilePath(path, ExcelHelper.ExcelMode.READ);
		sheet = helper.getSheet("SafeRead");

		Object val = helper.readCell(sheet, 0, 0); // should be null, not crash
		assertNull(val);

		helper.closeWorkbook();
	}

	@Test
	@Order(30)
	void testWriteMultiSheetLargeData() throws IOException, InvalidFormatException {
		String path = filePath("MultiSheetWrite");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);

		List<List<Object>> data = new ArrayList<>();
		for (int i = 0; i < 100_000; i++) {
			data.add(Arrays.asList("Val" + i, i));
		}

		helper.writeMultiSheetLargeData(data, 25000, path, Arrays.asList("Name", "Index"));
		helper.closeWorkbook();
	}

	@Test
	@Order(31)
	void testWriteAndReadTextWithLeadingZeros() throws IOException {
		String path = filePath("LeadingZerosTest");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("ZeroTest");

		helper.writeCellAsString(sheet, 0, 0, "000123");
		helper.setCellValueAsText(sheet, 1, 0, "0456");

		helper.saveWorkbook(path);
		helper.closeWorkbook();

		helper.setFilePath(path, ExcelHelper.ExcelMode.READ);
		sheet = helper.getSheet("ZeroTest");

		assertEquals("000123", helper.readCellAsString(sheet, 0, 0));
		assertEquals("0456", helper.readCellAsString(sheet, 1, 0));

		helper.closeWorkbook();
	}

	@Test
	@Order(32)
	void testSetBordersWithIndividualStyles() throws IOException {
		String path = filePath("MultiBorderTest");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Borders");

		helper.writeCell(sheet, 0, 0, "Cell");
		Cell cell = sheet.getRow(0).getCell(0);
		helper.setBorders(cell, BorderStyle.THIN, BorderStyle.MEDIUM, BorderStyle.DASH_DOT, BorderStyle.DOTTED);

		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(33)
	void testCreateAndReadSheetWithSpecialChars() throws IOException {
		String path = filePath("SpecialSheetNames");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);

		String specialName = "Summary & Stats @2025";
		helper.createSheet(specialName);

		helper.saveWorkbook(path);
		helper.closeWorkbook();

		helper.setFilePath(path, ExcelHelper.ExcelMode.READ);
		assertNotNull(helper.getSheet(specialName));
		helper.closeWorkbook();
	}

	@Test
	@Order(34)
	void testReadRowThatDoesNotExist() throws IOException {
		String path = filePath("NonExistentRow");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		helper.createSheet("Empty");

		helper.saveWorkbook(path);
		helper.closeWorkbook();

		helper.setFilePath(path, ExcelHelper.ExcelMode.READ);
		Sheet sheet = helper.getSheet("Empty");
		assertNull(helper.readRow(sheet, 5)); // row does not exist
		helper.closeWorkbook();
	}

	@Test
	@Order(35)
	void testReadCellAsWrongType() throws IOException {
		String path = filePath("WrongTypeTest");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Types");

		helper.setCellValue(sheet, 0, 0, "NotABoolean");

		helper.saveWorkbook(path);
		helper.closeWorkbook();

		helper.setFilePath(path, ExcelHelper.ExcelMode.READ);
		sheet = helper.getSheet("Types");

		assertNull(helper.readCellAsBoolean(sheet, 0, 0)); // Should fail safely
		helper.closeWorkbook();
	}

	@Test
	@Order(36)
	void testMultipleSheetsWithSimilarNames() throws IOException {
		String path = filePath("SimilarNames");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);

		helper.createSheet("Report");
		helper.createSheet("Report_1");
		helper.createSheet("Report_2");

		List<String> names = helper.listSheetNames();
		assertTrue(names.contains("Report_1"));
		assertEquals(3, names.size());

		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(37)
	void testOverwriteExistingCellValue() throws IOException {
		String path = filePath("OverwriteTest");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);

		Sheet sheet = helper.createSheet("Overwrite");
		helper.writeCell(sheet, 0, 0, "Initial");
		helper.writeCell(sheet, 0, 0, "Updated");

		helper.saveWorkbook(path);
		helper.closeWorkbook();

		helper.setFilePath(path, ExcelHelper.ExcelMode.READ);
		sheet = helper.getSheet("Overwrite");

		assertEquals("Updated", helper.readCellAsString(sheet, 0, 0));
		helper.closeWorkbook();
	}

	@Test
	@Order(38)
	void testMergeMultipleRanges() throws IOException {
		String path = filePath("MultiMerge");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);

		Sheet sheet = helper.createSheet("Merge");

		helper.mergeCells(sheet, 0, 0, 0, 2); // A1:C1
		helper.mergeCells(sheet, 1, 2, 0, 0); // A2:A3

		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(39)
	void testWriteLargeDataInChunks() throws IOException {
		String path = filePath("ChunkWrite");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);

		Sheet sheet = helper.createSheet("Chunk");

		List<List<Object>> data = new ArrayList<>();
		for (int i = 0; i < 1000; i++) {
			data.add(Arrays.asList("Row" + i, i));
		}

		helper.writeLargeData(sheet, data.iterator(), 200); // Chunk size 200
		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(40)
	void testCellWithNullAndBlankChecks() throws IOException {
		String path = filePath("BlankNullCheck");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);

		Sheet sheet = helper.createSheet("BlankCheck");

		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setBlank();

		assertTrue(helper.isEmptyCell(cell));
		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(41)
	void generateAutomationTestReport() throws IOException {
		String path = filePath("AutomationTestResults");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet summary = helper.createSheet("Summary");

		// Header row with column titles
		List<Object> headers = Arrays.asList("Test ID", "Test Name", "Status", "Execution Time (ms)", "Link");
		helper.writeHeaders(summary, headers);
		helper.freezePane(summary, 1, 0);

		// Apply custom column widths
		int[] columnWidths = { 20, 40, 15, 25, 15 }; // in characters (approx.)
		for (int i = 0; i < columnWidths.length; i++) {
			summary.setColumnWidth(i, columnWidths[i] * 256); // Excel column width units
		}

		// Custom header styling (bold, center, font size, color)
		Row headerRow = summary.getRow(0);
		CellStyle headerStyle = summary.getWorkbook().createCellStyle();
		Font headerFont = summary.getWorkbook().createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 11);
		headerFont.setColor(IndexedColors.WHITE.getIndex());
		headerStyle.setFont(headerFont);
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		headerStyle.setWrapText(true);
		headerStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		for (int i = 0; i < headers.size(); i++) {
			Cell cell = headerRow.getCell(i);
			if (cell != null) {
				cell.setCellStyle(headerStyle);
			}
		}

		// Row data + styles
		for (int i = 1; i <= 100; i++) {
			String testId = "TC-" + String.format("%03d", i);
			String testName = "Test Case " + i;
			String status = i % 5 == 0 ? "FAIL" : "PASS";
			long execTime = (long) (Math.random() * 2000 + 200);

			String detailSheetName = "TC" + i;
			Sheet detailSheet = helper.createSheet(detailSheetName);
			helper.writeRow(detailSheet, 0, Arrays.asList("This is the detail section for test case " + testId));

			Row row = summary.createRow(i);
			row.setHeightInPoints(20);

			row.createCell(0).setCellValue(testId);
			row.createCell(1).setCellValue(testName);
			row.createCell(2).setCellValue(status);
			row.createCell(3).setCellValue(execTime);

			// Create hyperlink
			helper.createHyperlinkToCellInDifferentSheet(new ExcelHelper.CellRef(summary, i, 4), detailSheetName, 0, 0,
					"Details");

			// Basic font styling for row
			for (int j = 0; j < 5; j++) {
				Cell cell = row.getCell(j);
				if (cell == null)
					cell = row.createCell(j);

				CellStyle cellStyle = summary.getWorkbook().createCellStyle();
				Font font = summary.getWorkbook().createFont();
				font.setFontHeightInPoints((short) 10);
				cellStyle.setFont(font);
				cellStyle.setWrapText(true);
				cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

				// Zebra striping
				if (i % 2 == 0) {
					cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				}

				// Center align status
				if (j == 2 || j == 3) {
					cellStyle.setAlignment(HorizontalAlignment.CENTER);
				}

				cell.setCellStyle(cellStyle);
			}

			// Status-specific background
			Cell statusCell = row.getCell(2);
			IndexedColors statusColor = status.equals("PASS") ? IndexedColors.LIGHT_GREEN : IndexedColors.RED;
			helper.setCellBackgroundColor(statusCell, statusColor);
		}

		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(42)
	void generateSalesReport() throws IOException {
		String path = filePath("SalesReport");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Sales");

		helper.writeHeaders(sheet,
				Arrays.asList("Region", "Sales Rep", "Product", "Units", "Revenue", "Commission %", "Commission"));
		helper.applyHeaderStyle(sheet, 0);
		helper.freezePane(sheet, 1, 0);

		String[] regions = { "North", "South", "East", "West" };
		String[] products = { "Laptop", "Phone", "Monitor", "Keyboard" };

		for (int i = 1; i <= 100; i++) {
			String region = regions[i % 4];
			String rep = "Rep " + i;
			String product = products[i % 4];
			int units = (int) (Math.random() * 50 + 10);
			double revenue = units * (Math.random() * 100 + 50);
			double commissionPct = 0.05 + (Math.random() * 0.1);
			double commission = revenue * commissionPct;

			List<Object> row = Arrays.asList(region, rep, product, units, revenue, commissionPct, commission);
			helper.writeRow(sheet, i, row);

			helper.formatCellAsCurrency(sheet.getRow(i).getCell(4));
			helper.formatCellAsPercent(sheet.getRow(i).getCell(5), 2);
			helper.formatCellAsCurrency(sheet.getRow(i).getCell(6));
		}

		helper.autoSizeAllColumns(sheet);
		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(43)
	void generateDatabaseExport() throws IOException {
		String path = filePath("DatabaseExport");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Users");

		List<Object> headers = Arrays.asList("User ID", "Name", "Email", "Signup Date", "Active");
		helper.writeHeaders(sheet, headers);
		helper.applyHeaderStyle(sheet, 0);
		helper.freezePane(sheet, 1, 0);

		for (int i = 1; i <= 100; i++) {
			String email = "user" + i + "@example.com";
			Date signup = new Date(System.currentTimeMillis() - (long) (Math.random() * 31536000000L)); // up to 1 year
																										// ago

			List<Object> row = Arrays.asList(i, "User " + i, email, signup, i % 2 == 0);
			helper.writeRow(sheet, i, row);
			helper.formatCellAsDate(sheet.getRow(i).getCell(3), "yyyy-MM-dd");
		}

		helper.autoSizeAllColumns(sheet);
		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(44)
	void generateMonitoringDashboard() throws IOException {
		String path = filePath("MonitoringDashboard");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Dashboard");

		helper.writeHeaders(sheet, Arrays.asList("Service", "Status", "CPU (%)", "RAM (MB)", "Last Check"));
		helper.applyHeaderStyle(sheet, 0);
		helper.freezePane(sheet, 1, 0);

		String[] services = { "Auth", "DB", "API", "Frontend", "Worker" };

		for (int i = 1; i <= 50; i++) {
			String service = services[i % services.length];
			String status = i % 7 == 0 ? "DOWN" : "UP";
			double cpu = Math.random() * 100;
			int ram = (int) (Math.random() * 8000 + 512);
			Date lastCheck = new Date();

			List<Object> row = Arrays.asList(service, status, cpu, ram, lastCheck);
			helper.writeRow(sheet, i, row);

			// Format
			helper.formatCellAsNumber(sheet.getRow(i).getCell(2), 1);
			helper.formatCellAsDate(sheet.getRow(i).getCell(4), "HH:mm:ss");

			// Color status
			Cell statusCell = sheet.getRow(i).getCell(1);
			helper.setCellBackgroundColor(statusCell,
					status.equals("UP") ? IndexedColors.LIGHT_GREEN : IndexedColors.RED);
		}

		helper.autoSizeAllColumns(sheet);
		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(45)
	void generateEmployeeDirectory() throws IOException {
		String path = filePath("EmployeeDirectory");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Directory");

		helper.writeHeaders(sheet, Arrays.asList("Emp ID", "Name", "Title", "Phone", "Email", "Profile Link"));
		helper.applyHeaderStyle(sheet, 0);
		helper.freezePane(sheet, 1, 0);
		helper.applyAutoFilter(sheet, 0);

		for (int i = 1; i <= 100; i++) {
			String name = "Employee " + i;
			String title = "Engineer Level " + (i % 5 + 1);
			String phone = "+1-555-01" + String.format("%03d", i);
			String email = "employee" + i + "@company.com";

			Row row = sheet.createRow(i);
			row.createCell(0).setCellValue("EMP" + i);
			row.createCell(1).setCellValue(name);
			row.createCell(2).setCellValue(title);
			row.createCell(3).setCellValue(phone);
			row.createCell(4).setCellValue(email);

			helper.createHyperlinkToURL(new ExcelHelper.CellRef(sheet, i, 5), "https://company.com/profiles/" + i,
					"View Profile");
		}

		helper.autoSizeAllColumns(sheet);
		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

	@Test
	@Order(46)
	void generateComparativeTaxSummaryReport() throws IOException {
		String path = filePath("ComparativeTaxSummaryReport");
		ExcelHelper helper = new ExcelHelper();
		helper.setFilePath(path, ExcelHelper.ExcelMode.WRITE);
		Sheet sheet = helper.createSheet("Tax Summary");

		// Row 0 - Title
		helper.writeRow(sheet, 0, Arrays.asList("Firm, Bank, Business or Company Name"));
		helper.mergeCells(sheet, 0, 0, 0, 13); // Merging across the full title area
		helper.setCellBackgroundColor(sheet.getRow(0).getCell(0), IndexedColors.YELLOW);
		helper.setBold(sheet.getRow(0).getCell(0));
		sheet.getRow(0).getCell(0).getCellStyle().setAlignment(HorizontalAlignment.CENTER);

		// Row 1 - Empty for spacing

		// Row 2 - Report Title
		helper.writeRow(sheet, 2, Arrays.asList("COMPARATIVE TAX  SUMMARY REPORT"));
		helper.mergeCells(sheet, 2, 2, 0, 13);
		helper.setBold(sheet.getRow(2).getCell(0));
		sheet.getRow(2).getCell(0).getCellStyle().setAlignment(HorizontalAlignment.CENTER);

		// Row 3 - Date Range
		helper.writeRow(sheet, 3, Arrays.asList("FROM 01/JULY/2021  TO 30/OCTOBER/2022"));
		helper.mergeCells(sheet, 3, 3, 0, 13);
		helper.setBold(sheet.getRow(3).getCell(0));
		sheet.getRow(3).getCell(0).getCellStyle().setAlignment(HorizontalAlignment.CENTER);

		// Row 4 - Date Label
		Row dateRow = sheet.createRow(4);
		Cell dateCell = dateRow.createCell(13);
		dateCell.setCellValue("Date: " + new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
		sheet.setColumnWidth(13, 25 * 256);

		// Row 5 - Header
		List<Object> headers = new ArrayList<>();
		headers.add("S. NO");
		headers.add("HEAD OF ACCOUNT");
		String[] months = { "July", "Aug", "Sept", "Oct" };
		for (String month : months) {
			headers.add("Budgeted " + month);
			headers.add("Actual " + month);
		}
		helper.writeHeaders(sheet, headers);

		Row headerRow = sheet.getRow(5);
		if (headerRow == null)
			headerRow = sheet.createRow(5);

		for (int i = 2, monthIndex = 0; i < headers.size(); i += 2, monthIndex++) {
			Cell budgetCell = headerRow.getCell(i);
			if (budgetCell == null)
				budgetCell = headerRow.createCell(i);
			Cell actualCell = headerRow.getCell(i + 1);
			if (actualCell == null)
				actualCell = headerRow.createCell(i + 1);

			IndexedColors budgetColor, actualColor;
			switch (monthIndex) {
			case 0 -> {
				budgetColor = IndexedColors.LIGHT_BLUE;
				actualColor = IndexedColors.BLUE;
			}
			case 1 -> {
				budgetColor = IndexedColors.TAN;
				actualColor = IndexedColors.BROWN;
			}
			case 2 -> {
				budgetColor = IndexedColors.LIGHT_ORANGE;
				actualColor = IndexedColors.ORANGE;
			}
			case 3 -> {
				budgetColor = IndexedColors.LIGHT_GREEN;
				actualColor = IndexedColors.GREEN;
			}
			default -> {
				budgetColor = IndexedColors.GREY_25_PERCENT;
				actualColor = IndexedColors.GREY_40_PERCENT;
			}
			}

			helper.setCellBackgroundColor(budgetCell, budgetColor);
			helper.setCellBackgroundColor(actualCell, actualColor);
		}

		// Sample Head of Account Rows
		String[] accounts = { "TOTAL", "IMPORT", "EXEMPTED", "BTL", "UN-PAID", "TAXABLE", "RATE", "TAXCALCULATED",
				"TAXPAID" };

		for (int i = 0; i < accounts.length; i++) {
			List<Object> row = new ArrayList<>();
			row.add(i + 1);
			row.add(accounts[i]);

			for (int j = 0; j < 4; j++) {
				row.add((int) (Math.random() * 10000)); // Budgeted
				row.add((int) (Math.random() * 10000)); // Actual
			}

			helper.writeRow(sheet, 6 + i, row);
		}

		helper.autoSizeAllColumns(sheet);
		helper.saveWorkbook(path);
		helper.closeWorkbook();
	}

}
