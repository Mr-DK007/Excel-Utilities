package com.excelUtilities;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.*;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.function.Predicate;
import java.util.stream.IntStream;

/**
 * A helper class for simplifying interactions with Excel files (.xls and
 * .xlsx). Supports large datasets (up to 3 million rows) and provides efficient
 * read/write operations. Includes basic colorful console logging.
 *
 * @author Mr.DK (Deepak)
 * @version 1.0
 * @since 2024-07-30
 */
public class ExcelHelper {

	Workbook workbook;
	private String filePath;
	private ExcelMode mode;

	// ANSI color codes
	private static final String RESET = "\u001B[0m";
	private static final String RED = "\u001B[31m";
	private static final String GREEN = "\u001B[32m";
	private static final String YELLOW = "\u001B[33m";
	private static final String BLUE = "\u001B[34m";

	private final Map<String, LogLevel> colorToLogLevelMap = new HashMap<>();

	/**
	 * Excel operating modes.
	 */
	public enum ExcelMode {
		READ, WRITE, STREAM
	}

	/**
	 * Represents supported Excel file types.
	 */
	public enum ExcelFileType {
		XLS, XLSX
	}

	/**
	 * Represents log level.
	 */
	public enum LogLevel { // Make public
		INFO, SUCCESS, WARNING, ERROR
	}

	/**
	 * Preset format options to simplify user styling. Acts like a dropdown to
	 * reduce formatting complexity for end users.
	 */
	public enum FormatPreset {
		CURRENCY("$#,##0.00"), PERCENT("0.0%"), DATE("yyyy-MM-dd"), TEXT("@"), NUMBER_2_DECIMAL("#,##0.00");

		private final String pattern;

		FormatPreset(String pattern) {
			this.pattern = pattern;
		}

		public String getPattern() {
			return pattern;
		}
	}

	/**
	 * Logs a formatted message to the console with dynamically determined log level
	 * and color.
	 * 
	 * @param message The message to log.
	 * @param color   The ANSI color code for the message, which also determines the
	 *                log level. Example usage: log("File opened.", GREEN);
	 *                //Success log("Invalid input.", RED); // Error
	 */

	private void log(String message, String color) {
		LogLevel level = colorToLogLevelMap.getOrDefault(color, LogLevel.INFO);

		// Format the message with padding to align the output:
		String formattedMessage = String.format("[%-10s] %s", level, message); // Pad the level to 10 characters

		// Apply color only to the log level:
		String coloredLevel = String.format("%s%s%s", color, level, RESET);
		String coloredMessage = formattedMessage.replace(level.toString(), coloredLevel); // Apply to only the log level

		System.out.println(coloredMessage);

	}

	public ExcelHelper() {
	}

	/**
	 * Sets file path and mode, then initializes workbook.
	 */
	public void setFilePath(String filePath, ExcelMode mode) throws IOException {
		this.filePath = filePath;
		this.mode = mode;
		initializeWorkbook();

		colorToLogLevelMap.put(BLUE, LogLevel.INFO);
		colorToLogLevelMap.put(GREEN, LogLevel.SUCCESS);
		colorToLogLevelMap.put(RED, LogLevel.ERROR);
		colorToLogLevelMap.put(YELLOW, LogLevel.WARNING);
	}

	/**
	 * Creates workbook instance based on file type and mode.
	 */
	public void createWorkbook() throws IOException {
		initializeWorkbook();
	}

	/**
	 * Initializes the workbook based on the mode and file extension. Uses
	 * SXSSFWorkbook when in STREAM mode or for .xlsx in WRITE mode with streaming
	 * enabled.
	 *
	 * @throws IOException if an I/O error occurs
	 */
	private void initializeWorkbook() throws IOException {
		if (mode == ExcelMode.WRITE || mode == ExcelMode.STREAM) {
			if (filePath.toLowerCase().endsWith(".xlsx") || mode == ExcelMode.STREAM) {
				workbook = new SXSSFWorkbook();
				log("Creating .xlsx workbook (streaming) at: " + filePath, GREEN);
			} else if (filePath.toLowerCase().endsWith(".xls")) {
				workbook = new HSSFWorkbook();
				log("Creating .xls workbook at: " + filePath, GREEN);
			} else {
				throw new IllegalArgumentException("File type must be .xls or .xlsx");
			}
		} else if (mode == ExcelMode.READ) {
			try (FileInputStream fis = new FileInputStream(filePath)) {
				if (filePath.toLowerCase().endsWith(".xlsx")) {
					workbook = new XSSFWorkbook(fis);
					log("Opening .xlsx workbook in READ mode at: " + filePath, GREEN);
				} else if (filePath.toLowerCase().endsWith(".xls")) {
					workbook = new HSSFWorkbook(fis);
					log("Opening .xls workbook in READ mode at: " + filePath, GREEN);
				} else {
					throw new IllegalArgumentException("File type must be .xls or .xlsx");
				}
			}
		}
	}

	/**
	 * Initializes a streaming workbook (SXSSFWorkbook) for writing large XLSX files
	 * efficiently. Keeps a limited number of rows in memory to reduce memory usage.
	 *
	 * @return SXSSFWorkbook instance
	 *
	 * @example SXSSFWorkbook workbook = createStreamingWorkbook();
	 */
	public SXSSFWorkbook createStreamingWorkbook() {
		SXSSFWorkbook streamingWorkbook = new SXSSFWorkbook(100);
		streamingWorkbook.setCompressTempFiles(true);
		log("Initialized SXSSFWorkbook for streaming large data sets.", GREEN);
		return streamingWorkbook;
	}

	/**
	 * Writes a large dataset to the specified sheet using parallel processing.
	 * NOTE: Thread-safe only at row level. Avoid using styles inside this method
	 * unless properly synchronized.
	 *
	 * @param sheet The sheet to write to.
	 * @param data  The list of row data to write.
	 *
	 * @example excelHelper.writeDataParallel(sheet, dataList);
	 */
	public void writeDataParallel(Sheet sheet, List<List<Object>> data) {
		IntStream.range(0, data.size()).parallel().forEach(i -> {
			synchronized (sheet) {
				writeRow(sheet, i, data.get(i));
			}
		});
		log("Parallel data writing completed on sheet: " + sheet.getSheetName(), GREEN);
	}

	/**
	 * Validates values in a specific column using the provided predicate.
	 *
	 * @param sheet     The sheet to validate.
	 * @param colIndex  The column index to validate.
	 * @param validator A predicate to apply to each value.
	 * @return True if all values pass validation, false if any value fails.
	 *
	 * @example boolean valid = excelHelper.validateColumnValues(sheet, 0, val ->
	 *          val instanceof String);
	 */
	public boolean validateColumnValues(Sheet sheet, int colIndex, Predicate<Object> validator) {
		for (Row row : sheet) {
			Cell cell = row.getCell(colIndex);
			Object value = getCellValue(cell);
			if (!validator.test(value)) {
				log(String.format("Validation failed at row %d, column %d", row.getRowNum(), colIndex), RED);
				return false;
			}
		}
		log("All values in column validated successfully.", GREEN);
		return true;
	}

	/**
	 * Applies a predefined format to a cell (e.g., currency, percent). This makes
	 * it easier for users to select from dropdown-style formatting presets.
	 *
	 * @param cell   The cell to format.
	 * @param preset The formatting preset to apply.
	 *
	 * @example excelHelper.applyFormatPreset(cell, FormatPreset.CURRENCY);
	 */
	public void applyFormatPreset(Cell cell, FormatPreset preset) {
		CellStyle style = workbook.createCellStyle();
		style.setDataFormat(workbook.createDataFormat().getFormat(preset.getPattern()));
		cell.setCellStyle(style);
		log(String.format("Applied format '%s' to cell in sheet: %s", preset.name(), cell.getSheet().getSheetName()),
				GREEN);
	}

	private Object getCellValue(Cell cell) {
		if (cell == null)
			return null;
		switch (cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue();
		case NUMERIC:
			return DateUtil.isCellDateFormatted(cell) ? cell.getDateCellValue() : cell.getNumericCellValue();
		case BOOLEAN:
			return cell.getBooleanCellValue();
		case FORMULA:
			return cell.getCellFormula();
		default:
			return null;
		}
	}

	/**
	 * Creates a new sheet in the workbook.
	 *
	 * @param sheetName The name of the sheet to create.
	 * @return The newly created sheet.
	 * @throws IllegalArgumentException If a sheet with the same name already
	 *                                  exists.
	 */

	public Sheet createSheet(String sheetName) {
		if (workbook.getSheet(sheetName) != null) {
			log(String.format("Sheet with name '%s' already exists.", sheetName), YELLOW); // Warning in yellow
			throw new IllegalArgumentException("Sheet with name '" + sheetName + "' already exists.");
		}
		log(String.format("Creating new sheet: %s", sheetName), GREEN);
		return workbook.createSheet(sheetName);
	}

	/**
	 * Writes a string value to a specific cell, preserving leading zeros.
	 *
	 * @param sheet  The sheet to write to.
	 * @param rowNum The row number (0-indexed).
	 * @param colNum The column number (0-indexed).
	 * @param value  The string value to write. Example: writeCellAsString(sheet, 0,
	 *               0, "00123"); // Writes "00123" to cell A1.
	 */
	public void writeCellAsString(Sheet sheet, int rowNum, int colNum, String value) {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			row = sheet.createRow(rowNum);
		}

		Cell cell = row.createCell(colNum);
		cell.setCellValue(value); // Sets as string
		log(String.format("Writing data '%s' to sheet '%s' at cell (%d, %d)", value, sheet.getSheetName(), rowNum,
				colNum), GREEN);

	}

	/**
	 * Gets a sheet by its name.
	 *
	 * @param sheetName The name of the sheet to retrieve.
	 * @return The sheet with the specified name, or null if not found. Example
	 *         usage: Sheet sheet1 = excelHelper.getSheet("Sheet1")
	 */

	public Sheet getSheet(String sheetName) {
		Sheet sheet = workbook.getSheet(sheetName);
		if (sheet == null) {
			log(String.format("Sheet '%s' not found.", sheetName), YELLOW);
		} else {
			log(String.format("Retrieving sheet '%s'", sheetName), GREEN);

		}
		return sheet;

	}

	/**
	 * Gets a sheet by its index.
	 *
	 * @param sheetIndex The index of the sheet to retrieve (0-based).
	 * @return The sheet at the specified index.
	 * @throws IllegalArgumentException If the sheet index is out of range. Example
	 *                                  usage: Sheet sheet1 =
	 *                                  excelHelper.getSheet(0) // Returns the first
	 *                                  sheet
	 */

	public Sheet getSheet(int sheetIndex) {
		try {
			Sheet sheet = workbook.getSheetAt(sheetIndex);
			log(String.format("Retrieving sheet at index %d : %s", sheetIndex, sheet.getSheetName()), GREEN);
			return sheet;

		} catch (IllegalArgumentException e) {
			log(String.format("Invalid Sheet Index : %s", e.getMessage()), RED);
			throw new IllegalArgumentException("Sheet index out of range: " + sheetIndex, e);

		}
	}

	/**
	 * Writes a list of objects as a row in the specified sheet.
	 *
	 * @param sheet  The sheet to write to.
	 * @param rowNum The row number (0-indexed).
	 * @param values The list of values to write. Example Usage: writeRow(sheet, 2,
	 *               Arrays.asList("Value1", 123, true, 3.14));
	 */

	public void writeRow(Sheet sheet, int rowNum, List<Object> values) {

		Row row = sheet.getRow(rowNum);

		if (row == null) {
			row = sheet.createRow(rowNum);
		}
		int colNum = 0;
		for (Object value : values) {
			Cell cell = row.createCell(colNum++);

			if (value instanceof String) {
				cell.setCellValue((String) value);
			} else if (value instanceof Number) {
				cell.setCellValue(((Number) value).doubleValue());
			} else if (value instanceof Boolean) {
				cell.setCellValue((Boolean) value);
			} else if (value instanceof java.util.Date) {
				cell.setCellValue((java.util.Date) value);
			} else {
				cell.setCellValue(value.toString()); // Handle other data types as needed
			}

		}
		log(String.format("Writing Row to sheet '%s', at row index '%s'", sheet.getSheetName(), rowNum), GREEN);

	}

	/**
	 * Writes a generic value to a specific cell, automatically handling different
	 * data types. This method now preserves leading zeros in numbers when they are
	 * passed as strings.
	 *
	 * @param sheet  The sheet to write to.
	 * @param rowNum The row number (0-indexed).
	 * @param colNum The column index (0-indexed).
	 * @param value  The value to write to the cell.
	 */
	public void setCellValue(Sheet sheet, int rowNum, int colNum, Object value) {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			row = sheet.createRow(rowNum);
		}
		Cell cell = row.createCell(colNum);

		setCellValue(cell, value); // Helper method to handle cell values and preserve leading zeros.

		log(String.format("Writing value '%s' to cell (%d, %d) in sheet: %s", value, rowNum, colNum,
				sheet.getSheetName()), GREEN);

	}

	/**
	 * A helper method to set cell values, preserving leading zeros if the value is
	 * a number passed as a string.
	 * 
	 * @param cell  The cell to set the value for.
	 * @param value The value to set.
	 */
	private void setCellValue(Cell cell, Object value) {

		// Preserve leading zeros if passed as a String representing a number:
		if (value instanceof String) {
			String strVal = (String) value;
			if (strVal.matches("-?\\d+(\\.\\d+)?")) { // Check if value is numeric but represented by a string (may have
														// leading zeroes)
				try {
					Double num = Double.parseDouble(strVal);
					cell.setCellValue(num);
					CellStyle style = workbook.createCellStyle();
					style.setDataFormat(workbook.createDataFormat().getFormat("@")); // Format Cell as Text to preserve
																						// any formatting done by end
																						// user like leading zeroes
					cell.setCellStyle(style);

					return; // Exit to avoid setting as a string below again.
				} catch (NumberFormatException ex) { // Handle potential NumberFormatException during conversion, if any
					// If not a number set value as String
					cell.setCellValue(strVal);

					return;

				}

			}

			cell.setCellValue(strVal);

		}
		// Existing Logic
		else if (value instanceof Number) {
			cell.setCellValue(((Number) value).doubleValue());
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		} else if (value instanceof Date) {
			cell.setCellValue((Date) value);
		} else if (value instanceof Calendar) {
			cell.setCellValue((Calendar) value);

		} else {
			cell.setCellValue(value != null ? value.toString() : ""); // Handle other types
		}

	}

	/**
	 * Sets the value of a cell as text, regardless of the data type of the provided
	 * value. This method is useful when you need to preserve the exact formatting
	 * of the input, such as leading zeros in numbers or preventing automatic date
	 * formatting.
	 *
	 * @param sheet  The sheet containing the cell.
	 * @param rowNum The row index (0-based).
	 * @param colNum The column index (0-based).
	 * @param value  The value to set as text.
	 *
	 *               Example Usage: setCellValueAsText(sheet, 0, 0, 123); // Cell A1
	 *               will contain the text "123", not the number 123
	 *               setCellValueAsText(sheet, 0, 1, "00456"); // Leading zeros
	 *               preserved
	 */
	public void setCellValueAsText(Sheet sheet, int rowNum, int colNum, Object value) {

		Row row = sheet.getRow(rowNum);
		if (row == null) {
			row = sheet.createRow(rowNum); // Create new row if doesn't exist.

		}

		Cell cell = row.createCell(colNum);
		CellStyle style = workbook.createCellStyle(); // Create a new CellStyle
		style.setDataFormat(workbook.createDataFormat().getFormat("@")); // The "@" format ensures text is treated as a
																			// string
		cell.setCellStyle(style); // Apply the CellStyle to the cell before setting a value so the text format is
									// applied correctly.

		if (value != null) { // Handle null values
			cell.setCellValue(String.valueOf(value)); // Set the cell value as a string, if the value is not null

		} else {
			cell.setCellValue(""); // For Null or empty string handling you decide to set what you want to set
									// based on your requirements

		}

		log(String.format("Setting cell value as text: '%s' at (%d, %d) in sheet '%s'", value, rowNum, colNum,
				sheet.getSheetName()), GREEN);

	}

	/**
	 * Reads all data from the specified sheet and returns it as a list of lists.
	 *
	 * @param sheet The sheet to read from.
	 * @return A list of lists representing the sheet's data. Example usage :
	 *         List<List<Object>> data = excelHelper.readAllData(sheet)
	 */
	public List<List<Object>> readAllData(Sheet sheet) {
		List<List<Object>> data = new ArrayList<>();
		for (Row row : sheet) {
			List<Object> rowData = new ArrayList<>();
			for (Cell cell : row) {
				switch (cell.getCellType()) {
				case STRING:
					rowData.add(cell.getStringCellValue());
					break;
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						rowData.add(cell.getDateCellValue());
					} else {
						rowData.add(cell.getNumericCellValue());
					}
					break;
				case BOOLEAN:
					rowData.add(cell.getBooleanCellValue());
					break;
				default:
					rowData.add(null); // Or handle other cell types as needed
				}
			}
			data.add(rowData);
		}
		log(String.format("Reading data from Sheet: %s", sheet.getSheetName()), GREEN);
		return data;

	}

	/**
	 * Saves the workbook to the specified file path.
	 *
	 * @param filePath The path to save the workbook to. Can be a different path
	 *                 than what was used to open the file.
	 * @throws IOException If an I/O error occurs while saving. Example Usage:
	 *                     excelHelper.saveWorkBook("/Users/mac/Desktop/output.xlsx");
	 */

	public void saveWorkbook(String filePath) throws IOException {
		try (FileOutputStream fos = new FileOutputStream(filePath)) {
			workbook.write(fos);
			log(String.format("Workbook saved to: %s", filePath), GREEN);

		}
	}

	/**
	 * Closes the workbook and disposes SXSSF temp files if applicable. Should
	 * always be called to avoid memory leaks.
	 *
	 * @throws IOException if closing fails
	 *
	 * @example excelHelper.closeWorkbook();
	 */
	public void closeWorkbook() throws IOException {
		if (workbook != null) {
			workbook.close();
			if (workbook instanceof SXSSFWorkbook) {
				((SXSSFWorkbook) workbook).dispose();
				log("SXSSFWorkbook temp files disposed.", GREEN);
			}
			log("Workbook closed successfully.", GREEN);
		}
	}

	/**
	 * Writes a large dataset to the specified sheet using an iterator and batch
	 * processing. This method is optimized for handling large datasets efficiently
	 * by writing data in batches. Batch size is defaulted to 1000.
	 *
	 * @param sheet        The sheet to write to.
	 * @param dataIterator An iterator providing the data rows.
	 * @throws IOException If an I/O error occurs during writing.
	 *
	 *
	 *                     Example Usage: excelHelper.writeLargeData(sheet,
	 *                     dataIterator); //Default batchsize used which is 1000
	 */
	public void writeLargeData(Sheet sheet, Iterator<List<Object>> dataIterator) throws IOException {
		writeLargeData(sheet, dataIterator, 1000); // Default batch size is 1000

	}

	/**
	 * Writes a large dataset to the specified sheet using an iterator and batch
	 * processing. This method is optimized for handling large datasets efficiently
	 * by writing data in batches.
	 *
	 * @param sheet        The sheet to write to.
	 * @param dataIterator An iterator providing the data rows.
	 * @param batchSize    The number of rows to write in each batch.
	 * @throws IOException If an I/O error occurs during writing. Example Usage:
	 *                     excelHelper.writeLargeData(sheet, dataIterator, 5000);
	 */
	public void writeLargeData(Sheet sheet, Iterator<List<Object>> dataIterator, int batchSize) throws IOException {
		int rowIndex = sheet.getLastRowNum() + 1;

		while (dataIterator.hasNext()) {
			for (int i = 0; i < batchSize && dataIterator.hasNext(); i++) {
				List<Object> row = dataIterator.next();
				writeRow(sheet, rowIndex++, row);
			}

			// If using streaming workbook, flush to disk safely
			if (sheet instanceof SXSSFSheet sxssfSheet) {
				try {
					sxssfSheet.flushRows(batchSize);
				} catch (IOException e) {
					throw new RuntimeException("Error flushing streaming rows", e);
				}
			}
		}

	}

	/**
	 * Applies a default header style to the specified row in the given sheet.
	 *
	 * @param sheet     The sheet containing the header row.
	 * @param headerRow The row index (0-based) of the header row. Example Usage:
	 *                  applyHeaderStyle(sheet, 0); //Applies header style to the
	 *                  first row of given sheet
	 */
	public void applyHeaderStyle(Sheet sheet, int headerRow) {

		CellStyle headerStyle = workbook.createCellStyle(); // Create header style object
		Font headerFont = workbook.createFont(); // Create header font

		headerFont.setBold(true); // Make text bold
		headerStyle.setFont(headerFont); // set font to headerStyle
		headerStyle.setAlignment(HorizontalAlignment.CENTER); // Align the cell value at center

		Row row = sheet.getRow(headerRow);
		if (row != null) {
			for (Cell cell : row) {
				cell.setCellStyle(headerStyle); // Applies header style to cell

			}

		}
		log(String.format("Applied Header Style to the row '%s' of sheet: %s", headerRow, sheet.getSheetName()), GREEN);

	}

	/**
	 * Auto-sizes all columns in the given sheet for better readability.
	 *
	 * @param sheet The sheet to auto-size columns in. Example Usage:
	 *              excelHelper.autoSizeAllColumns(sheet);
	 */

	public void autoSizeAllColumns(Sheet sheet) {
		if (sheet instanceof SXSSFSheet) {
			((SXSSFSheet) sheet).trackAllColumnsForAutoSizing(); // ✅ Required for streaming
		}
		int maxColumns = 0;
		for (Row row : sheet) {
			if (row != null && row.getLastCellNum() > maxColumns) {
				maxColumns = row.getLastCellNum();
			}
		}
		for (int i = 0; i < maxColumns; i++) {
			sheet.autoSizeColumn(i);
		}
		log(String.format("Auto-sized all columns in sheet: %s", sheet.getSheetName()), GREEN);
	}

	/**
	 * Merges cells in the specified range within the given sheet.
	 *
	 * @param sheet    The sheet containing the cells to merge.
	 * @param firstRow Index of the first row (0-based) in the merge range.
	 * @param lastRow  Index of the last row (0-based) in the merge range.
	 * @param firstCol Index of the first column (0-based) in the merge range.
	 * @param lastCol  Index of the last column (0-based) in the merge range.
	 *                 Example usage: excelHelper.mergeCells(sheet, 0, 1, 0, 2)
	 *                 //Merges the cells from A1 to C2
	 *
	 */

	public void mergeCells(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		CellRangeAddress range = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol); // creates cell range where
																								// we have to perform
																								// merge
		sheet.addMergedRegion(range);
		log(String.format("Cells merged in sheet: %s, in range: %s", sheet.getSheetName(), range.formatAsString()),
				GREEN);

	}

	/**
	 * Sets the width of a specific column.
	 *
	 * @param sheet    The sheet containing the column.
	 * @param colIndex The index of the column to set the width for (0-based).
	 * @param width    The width to set, measured in character units. Example Usage:
	 *                 setColumnWidth(sheet, 0, 25); //Sets the width of the first
	 *                 column(A) to 25 characters
	 *
	 */

	public void setColumnWidth(Sheet sheet, int colIndex, int width) {

		sheet.setColumnWidth(colIndex, width * 256); // Width is set in units of 1/256th of a character width

		log(String.format("Set width '%s' to column index '%s' in Sheet: %s", width, colIndex, sheet.getSheetName()),
				GREEN);

	}

	/**
	 * Freezes panes in the specified sheet to keep rows or columns visible while
	 * scrolling.
	 *
	 * @param sheet   The sheet to freeze panes in.
	 * @param numRows The number of rows to freeze.
	 * @param numCols The number of columns to freeze. Example Usage:
	 *                freezePane(sheet, 1, 0); // Freezes the top row
	 */
	public void freezePane(Sheet sheet, int numRows, int numCols) {
		sheet.createFreezePane(numCols, numRows);
		log(String.format("Freezing %d rows, %d cols in sheet: %s", numRows, numCols, sheet.getSheetName()), GREEN);

	}

	/**
	 * Sets a background color for a cell.
	 *
	 * @param cell  The cell to set the background color for.
	 * @param color The IndexedColors enum representing the desired color. You can
	 *              replace it with other cell style color methods if you need more
	 *              flexibility (like using RGB, HSL etc.).
	 *
	 *              Example Usage: setCellBackgroundColor(cell,
	 *              IndexedColors.LIGHT_YELLOW)
	 *
	 *
	 */

	public void setCellBackgroundColor(Cell cell, IndexedColors color) {
		CellStyle style = workbook.createCellStyle(); // Creating cell style object
		style.setFillForegroundColor(color.getIndex()); // set background color by providing color name
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND); // fill the background with solid color
		cell.setCellStyle(style); // setting cell style

		log(String.format("Setting Background Color '%s' to Cell  in Sheet: %s", color.toString(),
				cell.getSheet().getSheetName()), GREEN);

	}

	/**
	 * Adds a comment to a specific cell.
	 *
	 * @param sheet       The sheet containing the cell.
	 * @param rowNum      The row index (0-based) of the cell.
	 * @param colNum      The column index (0-based) of the cell.
	 * @param commentText The text of the comment. Example usage :
	 *                    addCellComment(sheet, 0, 0, "This is a test") // Adds cell
	 *                    comment to A1
	 */
	public void addCellComment(Sheet sheet, int rowNum, int colNum, String commentText) {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			row = sheet.createRow(rowNum); // create a new row if row not found
		}

		Cell cell = row.getCell(colNum);
		if (cell == null) {
			cell = row.createCell(colNum); // Create new cell if cell not found
		}

		// When the comment box is visible, have it show in a 1x3 space
		ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
		Drawing<?> drawing = sheet.createDrawingPatriarch();
		Comment comment = drawing.createCellComment(anchor);

		RichTextString str = workbook.getCreationHelper().createRichTextString(commentText);
		comment.setString(str);

		cell.setCellComment(comment);
		log(String.format("Added Cell Comment '%s' at row %d and col %d in Sheet: %s", commentText, rowNum, colNum,
				sheet.getSheetName()), GREEN);

	}

	/**
	 * Formats a cell as a date with the given format.
	 *
	 * @param cell   The cell to format.
	 * @param format The date format string (e.g., "yyyy-MM-dd", "MM/dd/yyyy").
	 *               Example usage: formatCellAsDate(cell, "yyyy-MM-dd");
	 */
	public void formatCellAsDate(Cell cell, String format) {
		CellStyle dateStyle = workbook.createCellStyle();
		DataFormat dateFormat = workbook.createDataFormat();
		dateStyle.setDataFormat(dateFormat.getFormat(format));
		cell.setCellStyle(dateStyle);
		log(String.format("Formatting cell as date with format '%s' in sheet: %s", format,
				cell.getSheet().getSheetName()), GREEN);

	}

	/**
	 * Formats a cell as a number with the specified number of decimal places.
	 *
	 * @param cell          The cell to format.
	 * @param decimalPlaces The number of decimal places to display. Example Usage:
	 *                      formatCellAsNumber(cell, 2); //Formats with 2 decimals
	 *
	 *
	 */

	public void formatCellAsNumber(Cell cell, int decimalPlaces) {

		String format = "#."; // Building number format string
		for (int i = 0; i < decimalPlaces; i++) {
			format += "0";

		}

		CellStyle numberStyle = workbook.createCellStyle();
		DataFormat numberFormat = workbook.createDataFormat();

		numberStyle.setDataFormat(numberFormat.getFormat(format));
		cell.setCellStyle(numberStyle);
		log(String.format("Cell formatted as number with %d decimal places in sheet: %s", decimalPlaces,
				cell.getSheet().getSheetName()), GREEN);

	}

	/**
	 * Protects a sheet with or without a password.
	 *
	 * @param sheet    The sheet to protect.
	 * @param password The password to use for protection (can be null for no
	 *                 password).
	 *
	 *                 Example Usage: protectSheet(sheet,"password");
	 */

	public void protectSheet(Sheet sheet, String password) {
		sheet.protectSheet(password);
		if (password != null) {
			log(String.format("Protecting sheet: %s with password", sheet.getSheetName()), GREEN);
		} else {
			log(String.format("Protecting sheet: %s without password", sheet.getSheetName()), GREEN);
		}

	}

	/**
	 * Reads data from a sheet and returns it as a list of maps, where each map
	 * represents a row and maps column headers to cell values.
	 *
	 * @param sheet The sheet to read from.
	 * @return A list of maps representing the sheet's data. Example Usage:
	 *         List<Map<String, Object>> data = excelHelper.readDataAsMap(sheet);
	 */

	public List<Map<String, Object>> readDataAsMap(Sheet sheet) {

		List<Map<String, Object>> excelData = new ArrayList<>();
		Row headerRow = sheet.getRow(0);

		if (headerRow != null) {
			int numColumns = headerRow.getPhysicalNumberOfCells(); // get total number of physical cells

			for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Iterate through data rows from row index 1 onwards

				Row dataRow = sheet.getRow(i);

				if (dataRow != null) {

					Map<String, Object> rowMap = new HashMap<>();
					for (int j = 0; j < numColumns; j++) { // Iterate through each cell of the data row

						Cell headerCell = headerRow.getCell(j);
						Cell dataCell = dataRow.getCell(j);

						String headerName = headerCell != null ? headerCell.getStringCellValue() : String.valueOf(j); // Get
																														// header
																														// value
																														// or
																														// use
																														// index
																														// if
																														// the
																														// header
																														// is
																														// null
						Object cellValue = getCellValue(dataCell); // Helper method

						rowMap.put(headerName, cellValue);

					}

					excelData.add(rowMap);

				}
			}
		}
		log(String.format("Reading data as maps from sheet: %s", sheet.getSheetName()), GREEN);

		return excelData;

	}

	/**
	 * Opens an existing Excel workbook.
	 *
	 * @param filePath The path to the Excel file to open.
	 * @throws IOException              If an I/O error occurs.
	 * @throws IllegalArgumentException If the file type is invalid. Example Usage:
	 *                                  excelHelper.openWorkbook("/Users/mac/Documents/data.xlsx");
	 */
	public void openWorkbook(String filePath) throws IOException {

		try (FileInputStream fis = new FileInputStream(filePath)) {

			if (filePath.toLowerCase().endsWith(".xlsx")) {
				workbook = new XSSFWorkbook(fis);
				log(String.format("Opening existing .xlsx workbook: %s", filePath), GREEN); // Log success

			} else if (filePath.toLowerCase().endsWith(".xls")) {
				workbook = new HSSFWorkbook(fis);

				log(String.format("Opening existing .xls workbook: %s", filePath), GREEN); // Log Success
			} else {
				log(String.format("Invalid file type: %s", filePath), RED); // Log error

				throw new IllegalArgumentException("Invalid file type.  Must be .xls or .xlsx");
			}
		}

	}

	/**
	 * Removes a sheet from the workbook. This method is overloaded to accept either
	 * sheet name or index.
	 * 
	 * @param sheetName The name of the sheet to remove. Example usage:
	 *                  removeSheet("mySheet");
	 *
	 */

	public void removeSheet(String sheetName) {
		int sheetIndex = workbook.getSheetIndex(sheetName); // Find index of sheet with given name.

		if (sheetIndex != -1) { // If sheet exists, then remove it
			workbook.removeSheetAt(sheetIndex);
			log(String.format("Sheet '%s' removed.", sheetName), GREEN);
		} else {
			log(String.format("No sheet found with the name '%s'.", sheetName), YELLOW);

		}

	}

	/**
	 * Removes a sheet from the workbook.
	 *
	 * This method is overloaded to accept either sheet name or index.
	 * 
	 * @param sheetIndex The index of the sheet to remove (0-based). Example Usage:
	 *                   removeSheet(1);
	 */
	public void removeSheet(int sheetIndex) {
		try {
			workbook.removeSheetAt(sheetIndex);
			log(String.format("Sheet at index %d removed.", sheetIndex), GREEN); // log successful sheet removal

		} catch (IllegalArgumentException ex) {

			log(String.format("Invalid sheet index '%s', sheet not found.", sheetIndex), YELLOW);

		}

	}

	/**
	 * Lists the names of all sheets in the workbook.
	 *
	 * @return A list of strings containing the names of all sheets. Example usage:
	 *         List<String> sheetNames = excelHelper.listSheetNames();
	 *
	 */

	public List<String> listSheetNames() {
		List<String> sheetNames = new ArrayList<>();

		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			sheetNames.add(workbook.getSheetName(i));

		}
		log(String.format("Listing all sheet names: %s", sheetNames), GREEN);

		return sheetNames;
	}

	/**
	 * Writes a generic value to a specific cell, automatically handling different
	 * data types.
	 * 
	 * @param sheet  The sheet to write to.
	 * @param rowNum The row number (0-indexed).
	 * @param colNum The column number (0-indexed).
	 * @param value  The value to write to the cell. Example usage :
	 *               writeCell(sheet, rowIndex, colIndex, value);
	 */
	public void writeCell(Sheet sheet, int rowNum, int colNum, Object value) {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			row = sheet.createRow(rowNum);
		}
		Cell cell = row.createCell(colNum);

		if (value instanceof String) {
			cell.setCellValue((String) value);
		} else if (value instanceof Number) {
			cell.setCellValue(((Number) value).doubleValue());
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		} else if (value instanceof Date) { // Covers both java.util.Date and java.sql.Date
			cell.setCellValue((Date) value);
		} else if (value instanceof Calendar) {
			cell.setCellValue((Calendar) value);
		} else {
			cell.setCellValue(value != null ? value.toString() : ""); // Handle other data types or null as needed
		}
		log(String.format("Writing value '%s' to cell (%d, %d) in sheet: %s", value, rowNum, colNum,
				sheet.getSheetName()), GREEN);

	}

	/**
	 * Writes a list of lists of objects to a sheet, where each inner list
	 * represents a row.
	 *
	 * @param sheet The sheet to write to.
	 * @param data  The data to write, as a list of lists of objects. Example usage:
	 *              writeData(sheet, data);
	 */

	public void writeData(Sheet sheet, List<List<Object>> data) {

		int rowNum = 1;
		for (List<Object> rowData : data) {
			writeRow(sheet, rowNum++, rowData); // Reusing the writeRow method
		}
		log(String.format("Writing data to the sheet: %s", sheet.getSheetName()), GREEN);

	}

	/**
	 * Writes headers to a sheet.
	 *
	 * @param sheet   The sheet to write to.
	 * @param headers The header row as a list of objects. Example usage:
	 *                writeHeaders(sheet, Arrays.asList("Column 1", "Column 2",
	 *                "Column3"));
	 *
	 */

	public void writeHeaders(Sheet sheet, List<Object> headers) {
		writeRow(sheet, 0, headers); // Write headers to the first row (index 0)
		log(String.format("Writing headers to sheet: %s", sheet.getSheetName()), GREEN);
	}

	/**
	 * Reads a cell value as a string. Handles various data types and formats.
	 * 
	 * @param sheet  The sheet containing the cell.
	 * @param rowNum The row index (0-based).
	 * @param colNum The column index (0-based).
	 * @return The cell value as a string, or null if the cell is empty or null.
	 *         Example Usage: String cellValue = excelHelper.readCellAsString(sheet,
	 *         rowIndex, columnIndex);
	 */
	public String readCellAsString(Sheet sheet, int rowNum, int colNum) {

		Row row = sheet.getRow(rowNum); // Get the row index first

		if (row != null) {
			Cell cell = row.getCell(colNum); // Get the cell number

			if (cell != null) {

				DataFormatter formatter = new DataFormatter(); // Using POI data formatter to handle different data
																// types
				String value = formatter.formatCellValue(cell);

				log(String.format("Reading cell (%d, %d) as string: %s", rowNum, colNum, value), GREEN);
				return value;

			}

		}

		log(String.format("Cell at (%d, %d) is null or empty.", rowNum, colNum), YELLOW);

		return null; // or empty string you decide according to your requirement

	}

	/**
	 * Reads a row of data from the specified sheet and returns it as a list of
	 * objects.
	 *
	 * @param sheet  The sheet to read from.
	 * @param rowNum The row number (0-indexed).
	 * @return A list of objects representing the row data, or null if the row is
	 *         empty. Example Usage: List<Object> myRow = excelHelper.readRow(sheet,
	 *         rowIndex);
	 */

	public List<Object> readRow(Sheet sheet, int rowNum) {

		Row row = sheet.getRow(rowNum);

		if (row != null) {

			List<Object> rowData = new ArrayList<>();
			Iterator<Cell> cellIterator = row.cellIterator(); // efficient way to iterate through cells
			while (cellIterator.hasNext()) {

				rowData.add(getCellValue(cellIterator.next())); // Reuse getCellValue helper method
			}
			log(String.format("Reading Row at index %d", rowNum), GREEN);

			return rowData;

		}
		log(String.format("Row at %d not found.", rowNum), YELLOW);

		return null;

	}

	/**
	 * Writes large data to multiple sheets, automatically splitting the data when a
	 * sheet reaches the maximum row limit.
	 * 
	 * @param data            The large dataset to write as a List of Lists of
	 *                        Objects.
	 * @param maxRowsPerSheet The maximum number of data rows allowed per sheet
	 *                        (excluding header, if any).
	 * @param outputFilePath  The base file path for output. Sheet numbers will be
	 *                        appended to the file name.
	 * @param headers         Optional headers for the data. If provided, they will
	 *                        be written to each sheet. Example usage:
	 *                        writeMultiSheetLargeData(largeDataset, 10000,
	 *                        "output_data", headers);
	 *
	 * @throws IOException            if an I/O error occurs during writing
	 * @throws InvalidFormatException if the output file format is invalid
	 */

	public void writeMultiSheetLargeData(List<List<Object>> data, int maxRowsPerSheet, String outputFilePath,
			List<Object> headers) throws IOException, InvalidFormatException {

		int sheetNum = 1;
		int rowIndex = 0;

		// Extract the file extension
		String fileExt = FilenameUtils.getExtension(outputFilePath);

		while (rowIndex < data.size()) { // Looping through dataset

			// Creates a new sheet with an incremented sheet number, maxrows is subtracted
			// from the total size for headers
			String newSheetName = "Sheet " + sheetNum;
			Sheet sheet = createSheet(newSheetName);

			// Constructs a new file path for each sheet generated.
			String outputFileName = String.format("%s_%d.%s", FilenameUtils.removeExtension(outputFilePath), sheetNum,
					fileExt);

			int rowsToWrite = Math.min(maxRowsPerSheet, data.size() - rowIndex); // Calculates how many rows to write in
																					// each sheet by taking the minimum
																					// of either rows remaining or rows
																					// to write

			// Writes header row if available
			if (headers != null && !headers.isEmpty()) {
				writeRow(sheet, 0, headers);

			}

			for (int i = 0; i < rowsToWrite; i++) {

				if (headers != null && !headers.isEmpty()) { // Add +1 if header present
					writeRow(sheet, i + 1, data.get(rowIndex++));

				} else {

					writeRow(sheet, i, data.get(rowIndex++));

				}

			}

			// Save per-sheet data to avoid excessive memory usage.
			try (FileOutputStream fos = new FileOutputStream(outputFileName)) {

				workbook.write(fos); // Works for both SXSSF and XSSF

			}

			sheetNum++; // Increment the sheet number
			workbook = new XSSFWorkbook(); // Creates a new workbook

		}
		log(String.format("Large Dataset split and written into multiple Sheet Files. Output base path: %s",
				outputFilePath), GREEN);

	}

	/**
	 * Reads a cell as an integer. Returns null if the cell is empty, null, or not a
	 * numeric type.
	 *
	 * @param sheet  The sheet containing the cell.
	 * @param rowNum The row index (0-based).
	 * @param colNum The column index (0-based).
	 * @return The cell value as an Integer, or null if an error occurs. Example
	 *         Usage : int myValue = excelHelper.readCellAsInt(sheet, rowIndex,
	 *         colIndex);
	 *
	 */
	public Integer readCellAsInt(Sheet sheet, int rowNum, int colNum) {
		Object value = readCell(sheet, rowNum, colNum); // Uses generic readCell method to read the object first
		if (value instanceof Number) {
			log(String.format("Reading cell (%d, %d) as integer: %d", rowNum, colNum, ((Number) value).intValue()),
					GREEN);
			return ((Number) value).intValue();
		} else if (value == null) {
			log(String.format("Cell at (%d, %d) is empty or null.", rowNum, colNum), YELLOW);
			return null;
		} else {
			log(String.format("Cannot read cell (%d, %d) as integer. Value is not numeric: %s", rowNum, colNum,
					value.toString()), RED);
			return null; // Or throw an exception if you prefer.
		}
	}

	/**
	 * Reads a cell from the specified sheet. Handles different cell types.
	 *
	 * @param sheet  The sheet to read from.
	 * @param rowNum The row index (0-based).
	 * @param colNum The column index (0-based).
	 * @return The cell's value as an Object, or null if the cell is empty or null.
	 *         Example Usage: Object value = readCell(sheet, 0, 1);
	 */

	public Object readCell(Sheet sheet, int rowNum, int colNum) {

		Row row = sheet.getRow(rowNum);
		if (row != null) {

			Cell cell = row.getCell(colNum);

			return getCellValue(cell); // Helper method

		}
		log(String.format("Cell at (%d, %d) is empty or null.", rowNum, colNum), YELLOW);

		return null;

	}

	/**
	 * Formats a cell as currency.
	 *
	 * @param cell The cell to format. Example Usage: formatCellAsCurrency(cell);
	 */
	public void formatCellAsCurrency(Cell cell) {

		CellStyle currencyStyle = workbook.createCellStyle();
		DataFormat format = workbook.createDataFormat();
		currencyStyle.setDataFormat(format.getFormat("$#,##0.00")); // Standard currency format
		cell.setCellStyle(currencyStyle);
		log(String.format("Cell formatted as currency in sheet: %s", cell.getSheet().getSheetName()), GREEN);

	}

	/**
	 * Forces a cell to be formatted as text, preserving leading zeros and other
	 * text formatting.
	 * 
	 * @param cell The cell to format. Example usage: forceStringFormat(cell);
	 *
	 */

	public void forceStringFormat(Cell cell) {
		CellStyle textStyle = workbook.createCellStyle();
		textStyle.setDataFormat(workbook.createDataFormat().getFormat("@")); // The "@" format signifies text formatting

		cell.setCellStyle(textStyle);
		log(String.format("Force Cell to String Formatted in sheet: %s", cell.getSheet().getSheetName()), GREEN);

	}

	/**
	 * Gets the row count of the specified sheet.
	 * 
	 * @param sheet The sheet to check.
	 * @return The number of rows in the sheet, or 0 if null or empty. Example
	 *         Usage: int rowCount = getRowCount(sheet);
	 */
	public int getRowCount(Sheet sheet) {
		if (sheet != null) {

			int lastRowNum = sheet.getLastRowNum();
			log(String.format("Getting row count from sheet: %s. Row Count: %d", sheet.getSheetName(), lastRowNum + 1),
					GREEN);
			return lastRowNum + 1; // +1 since row indices are 0-based

		} else {
			log("Cannot get row count. The provided sheet is null", YELLOW);
			return 0;
		}

	}

	/**
	 * Checks if a sheet contains a specific value.
	 *
	 * @param sheet The sheet to search in.
	 * @param value The value to search for.
	 * @return True if the value is found, false otherwise. Example usage: if
	 *         (sheetContains(sheet, "search_term")) { // Perform some action ... }
	 *
	 */

	public boolean sheetContains(Sheet sheet, Object value) {

		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell != null && getCellValue(cell).equals(value)) { // Use the helper method to compare values with
																		// type safety
					log(String.format("Value '%s' found in sheet '%s'", value, sheet.getSheetName()), GREEN);

					return true;
				}
			}
		}
		log(String.format("Value '%s' not found in sheet '%s'", value, sheet.getSheetName()), YELLOW);

		return false;

	}

	/**
	 * Finds the first cell containing a specific value in the given sheet.
	 *
	 * @param sheet The sheet to search in.
	 * @param value The value to search for.
	 * @return The first cell containing the value, or null if not found. Example
	 *         Usage: Cell foundCell = excelHelper.findCell(sheet, "search_value");
	 */

	public Cell findCell(Sheet sheet, Object value) {

		for (Row row : sheet) { // Iterate through each row in the sheet
			for (Cell cell : row) { // Iterate through each cell in the row
				if (cell != null && getCellValue(cell).equals(value)) {
					log(String.format("Cell found with value '%s' in sheet '%s'", value, sheet.getSheetName()), GREEN);
					return cell; // Return the first matching cell
				}

			}

		}

		log(String.format("No cell found with the value '%s' in sheet '%s'", value, sheet.getSheetName()), YELLOW);
		return null;

	}

	/**
	 * Finds all cells containing a specific value within the given sheet.
	 *
	 * @param sheet The sheet to search.
	 * @param value The value to search for.
	 * @return A list of cells containing the value, or an empty list if not found.
	 *         Example Usage: List<Cell> foundCells =
	 *         excelHelper.findAllCells(sheet, "value");
	 *
	 */

	public List<Cell> findAllCells(Sheet sheet, Object value) {
		List<Cell> foundCells = new ArrayList<>(); // Create an empty list to store found cells
		for (Row row : sheet) { // Iterating through rows

			for (Cell cell : row) { // Iterate through cells

				if (cell != null && getCellValue(cell).equals(value)) { // Check cell and use getCellValue() for type
																		// safety

					foundCells.add(cell); // Adds the matching cell to the list of all matching cells.
				}
			}
		}
		if (foundCells.isEmpty()) {
			log(String.format("No cell found with value '%s' in sheet '%s'", value, sheet.getSheetName()), YELLOW);

		} else {
			log(String.format("Found %d cells with value '%s' in sheet '%s'", foundCells.size(), value,
					sheet.getSheetName()), GREEN);
		}

		return foundCells; // Return all found cells
	}

	/**
	 * Gets the index of a column given its letter representation (e.g., "A", "B",
	 * "AA").
	 *
	 * @param columnLetter The column letter (case-insensitive).
	 * @return The column index (0-based), or -1 if the input is invalid. Example
	 *         Usage: int columnIndex = getColumnIndex("AA") // Returns 26
	 *         corresponding to AA column
	 */
	public int getColumnIndex(String columnLetter) {
		try {
			int columnIndex = CellReference.convertColStringToIndex(columnLetter); // Using poi's built in method
			log(String.format("Column Index for %s: %d", columnLetter, columnIndex), GREEN);

			return columnIndex;
		} catch (IllegalArgumentException ex) {

			log(String.format("Invalid column letter: %s", columnLetter), RED);
			return -1;

		}

	}

	/**
	 * Automatically splits a large dataset and writes it to multiple Excel files
	 * (xlsx format) to avoid memory issues. Each output file will contain up to
	 * maxRowsPerSheet rows of data.
	 *
	 * @param data               The large dataset to write (List of Lists of
	 *                           Objects).
	 * @param maxRowsPerSheet    The maximum number of rows allowed per output Excel
	 *                           file.
	 * @param outputFileBaseName The base name for output files. A number will be
	 *                           appended (e.g., output_1.xlsx, output_2.xlsx).
	 * @param headers            Optional header row. If provided, it will be added
	 *                           to each output file.
	 * @throws IOException            If any I/O errors occur during file writing.
	 * @throws InvalidFormatException If the output file format is invalid Example
	 *                                Usage : autoSplitFileIfTooLarge(data, 10000,
	 *                                "large_data_output", headerRow);
	 */
	public void autoSplitFileIfTooLarge(List<List<Object>> data, int maxRowsPerSheet, String outputFileBaseName,
			List<Object> headers) throws IOException, InvalidFormatException {

		int fileNum = 1;
		int rowIndex = 0;

		while (rowIndex < data.size()) {
			// Construct the output file path with incrementing file number.
			String outputFileName = String.format("%s_%d.xlsx", outputFileBaseName, fileNum++);

			try (Workbook tempWorkbook = new XSSFWorkbook();
					FileOutputStream fos = new FileOutputStream(outputFileName)) {
				Sheet sheet = tempWorkbook.createSheet("Sheet1"); // Creating sheet

				// write headers
				int startingRow = 0;

				if (headers != null && !headers.isEmpty()) {
					writeRowToTemp(sheet, 0, headers, tempWorkbook); // Use helper to create cells in temp workbook
					startingRow = 1; // Start writing from row 1 as row 0 will have headers
				}

				int rowsToWrite = Math.min(maxRowsPerSheet, data.size() - rowIndex);

				for (int i = 0; i < rowsToWrite; i++) {
					writeRowToTemp(sheet, i + startingRow, data.get(rowIndex++), tempWorkbook); // Helper method

				}

				tempWorkbook.write(fos); // Write data to file

			} // Close Workbook and output stream implicitly

		}
		log(String.format("Large dataset split into multiple files. Base name: %s", outputFileBaseName), GREEN);

	}

	/**
	 * Helper Method for WriteLargeData and writeRow method to use the tempWorkbook
	 * object reference.
	 */
	private void writeRowToTemp(Sheet sheet, int rowNum, List<Object> values, Workbook tempWorkbook) {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			row = sheet.createRow(rowNum);
		}

		int colNum = 0;
		for (Object value : values) {
			Cell cell = row.createCell(colNum++); // Create cells using tempWorkbook

			if (value instanceof String) {
				cell.setCellValue((String) value);
			} else if (value instanceof Number) {
				cell.setCellValue(((Number) value).doubleValue());
			} else if (value instanceof Boolean) {
				cell.setCellValue((Boolean) value);
			} else if (value instanceof java.util.Date) { // Includes java.sql.Date
				cell.setCellValue((java.util.Date) value);

			} else if (value instanceof Calendar) {
				cell.setCellValue((Calendar) value);
			} else {
				cell.setCellValue(value.toString()); // Handle other types as needed
			}

		}
	}

	// ... (Other existing methods in ExcelHelper)

	/**
	 * Reads a boolean value from a cell.
	 *
	 * @param sheet  The sheet containing the cell.
	 * @param rowNum The row index (0-based).
	 * @param colNum The column index (0-based).
	 * @return The boolean value of the cell, or null if the cell is empty, null, or
	 *         not a boolean type. Example Usage: Boolean boolVal =
	 *         readCellAsBoolean(sheet, rowIndex, columnIndex);
	 */
	public Boolean readCellAsBoolean(Sheet sheet, int rowNum, int colNum) {
		Object value = readCell(sheet, rowNum, colNum);

		if (value instanceof Boolean) {
			log(String.format("Reading cell (%d, %d) as boolean: %b", rowNum, colNum, (Boolean) value), GREEN);
			return (Boolean) value;
		} else if (value == null) {
			log(String.format("Cell at (%d, %d) is empty or null.", rowNum, colNum), YELLOW);

			return null; // Or handle as needed
		} else {
			log(String.format("Cannot read cell (%d, %d) as boolean. Value is not boolean type: %s", rowNum, colNum,
					value), RED);

			return null; // Or throw an exception, etc.

		}

	}

	/**
	 * Returns the letter representation of a column index.
	 * 
	 * @param columnIndex index starting from zero
	 * @return Column Letter Example usage : String colLetter = getColumnLetter(1);
	 *         // Returns "B"
	 */

	public String getColumnLetter(int columnIndex) {

		return CellReference.convertNumToColString(columnIndex);
	}

	/**
	 * Checks if a row is empty (contains no values).
	 *
	 * @param row The row to check.
	 * @return True if the row is empty, false otherwise. Example Usage : if
	 *         (isRowEmpty(row)){ // Perform some action ... }
	 *
	 */
	public boolean isRowEmpty(Row row) {
		if (row == null) { // Null check
			log("Row is null", YELLOW);
			return true; // Or handle null cases as needed.
		}

		for (Cell cell : row) { // Iterate through each cell
			if (cell != null && cell.getCellType() != CellType.BLANK) {
				log(String.format("Row is not empty in Sheet: %s", cell.getSheet().getSheetName()), GREEN);
				return false; // If at least one cell has value then the row is not empty.
			}
		}
		log(String.format("Row is empty in Sheet: %s", row.getSheet().getSheetName()), YELLOW);

		return true; // All cells are blank or null
	}

	/**
	 * Writes a large dataset to a single Excel file, creating new sheets as needed
	 * when the maximum row limit per sheet is reached.
	 *
	 * @param data            The large dataset to write (List of Lists of Objects).
	 * @param maxRowsPerSheet The maximum number of data rows per sheet.
	 * @param outputFilePath  The path to the output Excel file.
	 * @param sheetNameBase   The base name for the sheets. Numbers will be appended
	 *                        for additional sheets.
	 * @param headers         Optional header row (List of Strings). If provided,
	 *                        will be written to each sheet.
	 * @throws IOException If any I/O errors occur during file writing.
	 */
	public void writeLargeDataSingleFile(List<List<Object>> data, int maxRowsPerSheet, String outputFilePath,
			String sheetNameBase, List<Object> headers) throws IOException {

		int sheetNum = 1;
		int rowIndex = 0;
		Sheet currentSheet = null;
		int currentSheetStartingRow = 0;

		try (Workbook tempWorkbook = new XSSFWorkbook(); FileOutputStream fos = new FileOutputStream(outputFilePath)) {

			while (rowIndex < data.size()) {
				if (currentSheet == null
						|| (currentSheet.getLastRowNum() - currentSheetStartingRow + 1 >= maxRowsPerSheet)) { // Check
																												// if a
																												// new
																												// sheet
																												// is
																												// needed
					String sheetName = sheetNameBase + (sheetNum == 1 ? "" : sheetNum); // Sheet naming convention
																						// "SheetName" or "SheetName2",
																						// etc.
					currentSheet = tempWorkbook.createSheet(sheetName); // Creates new sheet

					// Write headers if provided
					currentSheetStartingRow = 0;
					if (headers != null && !headers.isEmpty()) {
						writeRowToTemp(currentSheet, 0, headers, tempWorkbook);
						currentSheetStartingRow = 1; // Update the first row for actual data
					}

					sheetNum++; // Increments the sheet number, read to create new sheet in the next iteration.
				}

				// Write the actual data to the currentSheet
				int rowsToWrite = Math.min(
						maxRowsPerSheet - (currentSheet.getLastRowNum() - currentSheetStartingRow + 1),
						data.size() - rowIndex);
				for (int i = 0; i < rowsToWrite; i++) {
					writeRowToTemp(currentSheet, currentSheet.getLastRowNum() + 1, data.get(rowIndex++), tempWorkbook);

				}
			}

			tempWorkbook.write(fos); // Write all the data to single output file

		}

		log(String.format("Large dataset written to a single file with multiple sheets: %s", outputFilePath), GREEN);

	}

	/**
	 * Applies a border to a cell.
	 *
	 * @param cell        The cell to apply the border to.
	 * @param borderStyle The border style to apply (e.g., THIN, MEDIUM, etc.)
	 *                    Example Usage: setBorder(cell, BorderStyle.THIN);
	 */
	public void setBorder(Cell cell, BorderStyle borderStyle) {
		CellStyle style = cell.getCellStyle(); // Get the current cell style if any exist
		if (style == null) {

			style = workbook.createCellStyle(); // Or create a new one if none exists
		}

		style.setBorderTop(borderStyle); // Applying the same border style to all sides.
		style.setBorderBottom(borderStyle);
		style.setBorderLeft(borderStyle);
		style.setBorderRight(borderStyle);
		cell.setCellStyle(style);

		log(String.format("Border applied to cell in sheet: %s with style: %s", cell.getSheet().getSheetName(),
				borderStyle), GREEN);

	}

	/**
	 * Applies a specific style to a range of cells.
	 * 
	 * @param sheet    The sheet containing the range of cells.
	 * @param firstRow The first row index (0-based) of the range.
	 * @param lastRow  The last row index (0-based) of the range.
	 * @param firstCol The first column index (0-based) of the range.
	 * @param lastCol  The last column index (0-based) of the range.
	 * @param style    The cell style to apply.
	 *
	 *                 Example Usage: applyRangeStyle(sheet, 1, 5, 0, 2,
	 *                 headerStyle); // Apply the headerStyle from rows 2 to 6, and
	 *                 from columns A to C.
	 */
	public void applyRangeStyle(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, CellStyle style) {
		for (int rowNum = firstRow; rowNum <= lastRow; rowNum++) {
			Row row = sheet.getRow(rowNum);

			if (row == null) {
				row = sheet.createRow(rowNum); // Create row if not exist
			}
			for (int colNum = firstCol; colNum <= lastCol; colNum++) {

				Cell cell = row.getCell(colNum);
				if (cell == null) {
					cell = row.createCell(colNum); // Create cell if not exist
				}
				cell.setCellStyle(style); // Apply style

			}
		}

		log(String.format("Range style applied to sheet: %s, from row %d to %d and col %d to %d", sheet.getSheetName(),
				firstRow, lastRow, firstCol, lastCol), GREEN);

	}

	/**
	 * Sets a formula for a cell.
	 *
	 * @param sheet   The sheet containing the cell.
	 * @param rowNum  The row index (0-based).
	 * @param colNum  The column index (0-based).
	 * @param formula The formula string (e.g., "SUM(A1:A10)"). Example Usage:
	 *                setCellFormula(sheet, 0, 2, "SUM(A1:B1)"); //Set sum formula
	 *                in C1
	 */

	public void setCellFormula(Sheet sheet, int rowNum, int colNum, String formula) {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			row = sheet.createRow(rowNum); // Create new row if doesn't exist.

		}
		Cell cell = row.createCell(colNum);
		cell.setCellFormula(formula);
		log(String.format("Formula '%s' set for cell (%d, %d) in sheet: %s", formula, rowNum, colNum,
				sheet.getSheetName()), GREEN);

	}

	/**
	 * Exports data from a sheet to a CSV file.
	 *
	 * @param sheet       The sheet to export.
	 * @param csvFilePath The path to save the CSV file.
	 * @param separator   The column separator character.
	 * @throws IOException If an I/O error occurs.
	 *
	 *                     Example usage: exportToCSV(sheet, "data.csv", ',');
	 */
	public void exportToCSV(Sheet sheet, String csvFilePath, char separator) throws IOException {

		try (BufferedWriter writer = new BufferedWriter(new FileWriter(csvFilePath))) { // BufferedWriter will improve
																						// performance by writing data
																						// in buffer
			DataFormatter dataFormatter = new DataFormatter(); // Using data formatter to handle data type and
																// formatting issues

			for (Row row : sheet) {

				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); // Handling missing cells
					String cellValue = dataFormatter.formatCellValue(cell);

					if (cellValue == null) {
						cellValue = ""; // Replace null with empty string or handle as needed

					}

					writer.write(cellValue);

					if (i < row.getLastCellNum() - 1) { // Avoid adding separator at the end of each row
						writer.write(separator);
					}

				}
				writer.newLine(); // End of row

			}

			log(String.format("Sheet '%s' exported to CSV: %s", sheet.getSheetName(), csvFilePath), GREEN);

		}
	}

	/**
	 * Imports data from a CSV file to a specified sheet.
	 *
	 * @param csvFilePath The path to the CSV file.
	 * @param sheet       The sheet to import the data into.
	 * @param separator   The column separator in the CSV file.
	 * @throws IOException If an I/O error occurs. Example Usage :
	 *                     excelHelper.importFromCSV("data.csv", sheet, ',')
	 */
	public void importFromCSV(String csvFilePath, Sheet sheet, char separator) throws IOException {

		try (BufferedReader reader = new BufferedReader(new FileReader(csvFilePath))) { // Improve performance

			String line;
			int rowNum = 0;

			while ((line = reader.readLine()) != null) { // Reads the file line by line
				String[] values = line.split(String.valueOf(separator)); // Splits the data on separator provided
				Row row = sheet.createRow(rowNum++); // Creates new row

				for (int colNum = 0; colNum < values.length; colNum++) { // Adding data cell by cell
					Cell cell = row.createCell(colNum);
					String value = values[colNum].trim();

					try {
						double num = Double.parseDouble(value); // Convert to Double if its numeric
						cell.setCellValue(num);
					} catch (NumberFormatException ex) { // If not a number set value as String
						cell.setCellValue(value);
					}

				}
			}

			log(String.format("Data imported from CSV '%s' into sheet: %s", csvFilePath, sheet.getSheetName()), GREEN);

		}

	}

	/**
	 * Applies specified border styles to the given cell.
	 * 
	 * @param cell   The cell to apply borders to.
	 * @param top    The top border style to set.
	 * @param bottom The bottom border style to set.
	 * @param left   The left border style to set.
	 * @param right  The right border style to set.
	 *
	 *               Example usage: setBorders(cell, BorderStyle.THIN,
	 *               BorderStyle.MEDIUM, BorderStyle.DASHED, BorderStyle.DOTTED)
	 *
	 */

	public void setBorders(Cell cell, BorderStyle top, BorderStyle bottom, BorderStyle left, BorderStyle right) {
		CellStyle style = cell.getCellStyle();

		if (style == null) {
			style = workbook.createCellStyle();
		}

		style.setBorderTop(top);
		style.setBorderBottom(bottom);
		style.setBorderLeft(left);
		style.setBorderRight(right);

		cell.setCellStyle(style);
		log(String.format("Borders applied to cell in sheet '%s'", cell.getSheet().getSheetName()), GREEN);

	}

	/**
	 * Applies a filter to the specified sheet.
	 *
	 * @param sheet  The sheet to apply the filter to.
	 * @param rowNum The row number containing the header row for filtering or, if
	 *               you are not filtering on headers, the row number above which to
	 *               place the filter Example Usage: applyFilter(sheet, 0); // Apply
	 *               filter using the header row (row 1)
	 */
	public void applyAutoFilter(Sheet sheet, int rowNum) {

		int lastCol = sheet.getRow(rowNum).getLastCellNum() - 1; // Getting the last column index so that AutoFilter
																	// will work correctly.
		sheet.setAutoFilter(new CellRangeAddress(rowNum, rowNum, 0, lastCol)); // Apply auto-filter
		log(String.format("Auto-filter applied to sheet '%s' at row %d", sheet.getSheetName(), rowNum), GREEN);

	}

	/**
	 * Validates if the given cell is empty.
	 *
	 * @param cell The cell to validate.
	 * @return True if the cell is empty or null, false otherwise. Example usage: if
	 *         (isEmptyCell(cell)){ // Perform some action ... }
	 *
	 */

	public boolean isEmptyCell(Cell cell) {

		if (cell == null || cell.getCellType() == CellType.BLANK
				|| (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty())) {
			log(String.format("Cell is empty in sheet: %s",
					(cell != null) ? cell.getSheet().getSheetName() : "null cell"), YELLOW);

			return true; // Covers blank, empty string, and null cases

		}

		log(String.format("Cell is not empty in sheet: %s", cell.getSheet().getSheetName()), GREEN);

		return false;

	}

	/**
	 * Reads a cell value and attempts to convert it to the specified type.
	 *
	 * @param sheet  The sheet containing the cell.
	 * @param rowNum The row index (0-based).
	 * @param colNum The column index (0-based).
	 * @param type   The desired data type (e.g., String.class, Integer.class,
	 *               Date.class).
	 * @param <T>    The generic type parameter.
	 * @return The cell value converted to the specified type, or null if the cell
	 *         is empty, null, or cannot be converted. Example usage: Integer value
	 *         = excelHelper.readCellAsType(sheet, 0, 0, Integer.class);
	 */

	public <T> T readCellAsType(Sheet sheet, int rowNum, int colNum, Class<T> type) {
		Object value = readCell(sheet, rowNum, colNum); // Read the cell using generic readCell method

		if (value == null) {
			log(String.format("Null or empty cell at row '%s' and column '%s' in sheet '%s'.", rowNum, colNum,
					sheet.getSheetName()), YELLOW);
			return null;

		}

		if (type.isInstance(value)) { // Checks if value is already an instance of the requested type, if its then
										// return it without any conversion
			log(String.format("Reading cell at row '%s' and column '%s' of type %s in sheet '%s'", rowNum, colNum,
					type.getName(), sheet.getSheetName()), GREEN);
			return type.cast(value);

		}

		if (type == String.class) {
			return type.cast(String.valueOf(value)); // Converts the cell value to a String and casts it.

		} else if (type == Integer.class && value instanceof Number) {
			return type.cast(((Number) value).intValue()); // Converts the number into Integer

		} else if (type == Double.class && value instanceof Number) {

			return type.cast(((Number) value).doubleValue()); // Converts the Number into Double

		} else if (type == Boolean.class && value instanceof Boolean) {

			return type.cast(value); // Casting it to Boolean

		} else if (type == Date.class && value instanceof Date) {

			return type.cast(value); // Casting to Date Object

		} else if (type == Float.class && value instanceof Number) { // Checking if the value is instance of number and
																		// then casting to Float.
			return type.cast(((Number) value).floatValue());

		}

		log(String.format("Cannot convert cell at row '%s' and column '%s' to type %s in sheet '%s'.", rowNum, colNum,
				type.getName(), sheet.getSheetName()), RED);

		return null; // Or throw an exception to handle incompatible types if you prefer

	}

	/**
	 * Sets the font for a given cell.
	 * 
	 * @param cell     The cell to apply to font to.
	 * @param fontName The name of the font (e.g., "Arial", "Calibri").
	 * @param fontSize The size of the font in points. Example Usage: setFont(cell,
	 *                 "Arial", 12);
	 *
	 */

	public void setFont(Cell cell, String fontName, short fontSize) {

		CellStyle style = cell.getCellStyle();

		if (style == null) {
			style = workbook.createCellStyle();
		}

		Font font = workbook.createFont(); // Create font object to apply changes
		font.setFontName(fontName); // set Font name
		font.setFontHeightInPoints(fontSize); // Set font size
		style.setFont(font); // Applying font change to cell
		cell.setCellStyle(style); // Applying the updated style to cell
		log(String.format("Setting font '%s' with size %d to cell in sheet '%s'", fontName, fontSize,
				cell.getSheet().getSheetName()), GREEN);

	}

	/**
	 * Sets the text color for a cell.
	 *
	 * @param cell  The cell to apply the color to.
	 * @param color The IndexedColors enum representing the color. Example Usage:
	 *              setTextColor(cell, IndexedColors.RED);
	 *
	 *
	 */
	public void setTextColor(Cell cell, IndexedColors color) {
		CellStyle style = cell.getCellStyle();

		if (style == null) {
			style = workbook.createCellStyle();

		}

		Font font = workbook.createFont();
		font.setColor(color.getIndex());
		style.setFont(font);
		cell.setCellStyle(style);
		log(String.format("Setting text color %s to cell in sheet '%s'", color, cell.getSheet().getSheetName()), GREEN);

	}

	/**
	 * Sets the horizontal alignment for a cell.
	 *
	 * @param cell      The cell to apply the alignment to.
	 * @param alignment The HorizontalAlignment enum representing the desired
	 *                  alignment. Example Usage: setHorizontalAlignment(cell,
	 *                  HorizontalAlignment.CENTER);
	 *
	 */
	public void setHorizontalAlignment(Cell cell, HorizontalAlignment alignment) {
		CellStyle style = cell.getCellStyle();
		if (style == null) {
			style = workbook.createCellStyle();
		}
		style.setAlignment(alignment);
		cell.setCellStyle(style); // Setting updated style
		log(String.format("Setting horizontal alignment to '%s' for cell in sheet '%s'.", alignment.toString(),
				cell.getSheet().getSheetName()), GREEN);

	}

	/**
	 * Sets the cell value as a numeric value and applies a specific data format.
	 * Useful when you need specific number formatting that's not covered by other
	 * formatting methods.
	 *
	 * @param cell      The cell to apply the number format to.
	 * @param number    The numeric value to set.
	 * @param formatStr The data format string. E.g., "#,##0.00" for two decimal
	 *                  places, "0.0%" for percentages, "yyyy-MM-dd" for dates, etc.
	 *                  Example usage: excelHelper.setNumberFormat(cell, 1234.567,
	 *                  "#,##0.00");
	 *
	 *
	 */
	public void setNumberFormat(Cell cell, double number, String formatStr) {
		CellStyle style = cell.getCellStyle(); // Get existing cell style if any or null
		if (style == null) {
			style = workbook.createCellStyle(); // Create new style if none present

		}

		DataFormat format = workbook.createDataFormat();
		style.setDataFormat(format.getFormat(formatStr)); // Setting custom format
		cell.setCellStyle(style);

		cell.setCellValue(number); // Setting the numeric value into the cell
		log(String.format("Number format '%s' applied to cell in sheet: '%s'", formatStr,
				cell.getSheet().getSheetName()), GREEN);

	}

	/**
	 * Sets the vertical alignment of the cell using the provided VerticalAlignment.
	 * 
	 * @param cell      The cell for which vertical alignment should be set.
	 * @param alignment VerticalAlignment The vertical alignment setting. Example
	 *                  Usage: excelHelper.setVerticalAlignment(cell,
	 *                  VerticalAlignment.CENTER);
	 */

	public void setVerticalAlignment(Cell cell, VerticalAlignment alignment) {
		CellStyle style = cell.getCellStyle();
		if (style == null) {
			style = workbook.createCellStyle();
		}

		style.setVerticalAlignment(alignment);
		cell.setCellStyle(style); // Update cell style
		log(String.format("Vertical alignment set to '%s' for cell in sheet '%s'.", alignment.toString(),
				cell.getSheet().getSheetName()), GREEN);

	}

	/**
	 * Wraps the text within the specified cell.
	 * 
	 * @param cell The cell in which to wrap text. Example Usage: wrapText(cell);
	 *
	 */
	public void wrapText(Cell cell) {

		CellStyle style = cell.getCellStyle(); // Retrieves the cell style.
		if (style == null) {
			style = workbook.createCellStyle(); // Or create it if it doesn't exist.

		}

		style.setWrapText(true); // Wraps text in cell.
		cell.setCellStyle(style); // set updated style
		log(String.format("Wrapped the text in cell in sheet '%s'", cell.getSheet().getSheetName()), GREEN);

	}

	/**
	 * Sets the row height for a specific row in points.
	 * 
	 * @param sheet  The sheet containing the row.
	 * @param rowNum The index of the row (0-based).
	 * @param height The height of the row in points.
	 *
	 *               Example usage: setRowHeight(sheet, 2, 25.5f); // Set row 3's
	 *               height to 25.5 points.
	 */
	public void setRowHeight(Sheet sheet, int rowNum, float height) {

		Row row = sheet.getRow(rowNum); // Retrieve specified row.
		if (row == null) {
			row = sheet.createRow(rowNum); // Creates row if it doesn't exist.
		}

		row.setHeightInPoints(height); // Sets row height.
		log(String.format("Row height set to '%s' points for row '%s' in sheet '%s'", height, rowNum,
				sheet.getSheetName()), GREEN);

	}

	/**
	 * Indents the cell content by the specified number of spaces.
	 *
	 * @param cell        The cell in which to indent.
	 * @param indentLevel The number of spaces to indent the content by. Example
	 *                    Usage: indentCell(cell, 2); // Indents the content by two
	 *                    spaces.
	 */

	public void indentCell(Cell cell, short indentLevel) {

		CellStyle style = cell.getCellStyle(); // Get cellStyle for specified cell, create if null.
		if (style == null) {
			style = workbook.createCellStyle();
		}
		style.setIndention(indentLevel); // Set indentation level for cell style.
		cell.setCellStyle(style); // set updated cell style
		log(String.format("Indented the cell at level %d in sheet '%s'", indentLevel, cell.getSheet().getSheetName()),
				GREEN);

	}

	/**
	 * Returns the number of columns (cells) in the specified row of a sheet. This
	 * method has been fixed to return a consistent 1-based count like POI's built
	 * in methods do. If you need to exclude empty or blank cells, then you must
	 * iterate and check for blank cell types or empty values and then subtract
	 * those from returned count based on your requirement.
	 * 
	 * 
	 * @param sheet    The sheet to check.
	 * @param rowIndex The index of the row.
	 * @return The number of cells in the specified row (1-based), or 0 if the row
	 *         is empty. Returns -1 if row not found.
	 */
	public int getColumnCount(Sheet sheet, int rowIndex) {
		if (sheet != null) {
			Row row = sheet.getRow(rowIndex); // Get specified row

			if (row != null) {
				// Get last cell num (1-based index, unlike 0-based row and cell indices)
				return row.getLastCellNum(); // Correctly handles empty rows

			} else {

				log(String.format("Row %d not found in sheet %s", rowIndex, sheet.getSheetName()), YELLOW);
				return -1; // Return -1 to indicate the row not found
			}
		} else {
			log("Sheet is null", YELLOW);

			return -1; // Return -1 if the sheet is null.

		}

	}

	/**
	 * Wrapper class to represent a cell reference in a sheet.
	 */
	public static class CellRef {
		public final Sheet sheet;
		public final int row;
		public final int col;

		public CellRef(Sheet sheet, int row, int col) {
			this.sheet = sheet;
			this.row = row;
			this.col = col;
		}

		public String getAddress() {
			return String.format("'%s'!%s%d", sheet.getSheetName(), getColumnLetter(col), row + 1);
		}

		private String getColumnLetter(int colIndex) {
			StringBuilder sb = new StringBuilder();
			while (colIndex >= 0) {
				sb.insert(0, (char) ('A' + (colIndex % 26)));
				colIndex = colIndex / 26 - 1;
			}
			return sb.toString();
		}

		public Sheet getSheet() {
			return sheet;
		}

		public int getRow() {
			return row;
		}

		public int getCol() {
			return col;
		}
	}

	/**
	 * Creates a hyperlink from a cell to a URL.
	 *
	 * @param source   The source cell reference.
	 * @param url      The URL to link to.
	 * @param linkText The display text for the hyperlink.
	 *
	 *                 Example: helper.createHyperlinkToURL(new CellRef(sheet, 0,
	 *                 0), "https://example.com", "Visit Site");
	 */
	public void createHyperlinkToURL(CellRef source, String url, String linkText) {
		createHyperlink(source.sheet, source.row, source.col, HyperlinkType.URL, url, linkText);
	}

	/**
	 * Creates a hyperlink from a cell to an email address.
	 *
	 * @param source       The source cell reference.
	 * @param emailAddress The email address.
	 * @param linkText     The display text for the hyperlink.
	 */
	public void createHyperlinkToEmail(CellRef source, String emailAddress, String linkText) {
		createHyperlink(source.sheet, source.row, source.col, HyperlinkType.EMAIL, "mailto:" + emailAddress, linkText);
	}

	/**
	 * Creates a hyperlink from a cell to a file path.
	 *
	 * @param source   The source cell reference.
	 * @param filePath The file path to link to.
	 * @param linkText The display text for the hyperlink.
	 */
	public void createHyperlinkToFile(CellRef source, String filePath, String linkText) {
		createHyperlink(source.sheet, source.row, source.col, HyperlinkType.FILE, "file:///" + filePath, linkText);
	}

	/**
	 * Creates a hyperlink from a cell to a cell range in another sheet.
	 *
	 * @param source    The source cell reference.
	 * @param destSheet The destination sheet.
	 * @param firstRow  The first row index (0-based).
	 * @param lastRow   The last row index (0-based).
	 * @param firstCol  The first column index (0-based).
	 * @param lastCol   The last column index (0-based).
	 * @param linkText  The display text for the hyperlink.
	 *
	 *                  Example: helper.createHyperlinkToRange(new CellRef(sheet1,
	 *                  0, 0), sheet2, 0, 1, 0, 1, "Go to Range");
	 */
	public void createHyperlinkToRange(CellRef source, Sheet destSheet, int firstRow, int lastRow, int firstCol,
			int lastCol, String linkText) {
		String address = String.format("'%s'!%s%d:%s%d", destSheet.getSheetName(), getColumnLetter(firstCol),
				firstRow + 1, getColumnLetter(lastCol), lastRow + 1);

		createHyperlink(source.sheet, source.row, source.col, HyperlinkType.DOCUMENT, address, linkText);
	}

	private void createHyperlink(Sheet sheet, int rowNum, int colNum, HyperlinkType type, String address,
			String linkText) {
		Row row = sheet.getRow(rowNum);
		if (row == null)
			row = sheet.createRow(rowNum);

		Cell cell = row.getCell(colNum);
		if (cell == null)
			cell = row.createCell(colNum);

		CreationHelper helper = workbook.getCreationHelper();
		Hyperlink link = helper.createHyperlink(type);
		link.setAddress(address);

		cell.setHyperlink(link);
		cell.setCellValue(linkText);

		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setUnderline(Font.U_SINGLE);
		font.setColor(IndexedColors.BLUE.getIndex());
		style.setFont(font);
		cell.setCellStyle(style);

		log(String.format("Hyperlink created in sheet '%s', cell (%d,%d), type '%s', address '%s': %s",
				sheet.getSheetName(), rowNum, colNum, type.toString(), address, linkText), GREEN);
	}

	/**
	 * Creates a hyperlink from one cell to a specific cell in a different sheet.
	 *
	 * @param fromCell        A reference to the source cell where the hyperlink
	 *                        will be placed.
	 * @param targetSheetName The name of the sheet to link to.
	 * @param targetRow       The zero-based row index of the target cell in the
	 *                        target sheet.
	 * @param targetCol       The zero-based column index of the target cell in the
	 *                        target sheet.
	 * @param linkText        The display text for the hyperlink in the source cell.
	 */
	public void createHyperlinkToCellInDifferentSheet(CellRef fromCell, String targetSheetName, int targetRow,
			int targetCol, String linkText) {
		Sheet sourceSheet = fromCell.getSheet();
		Row row = sourceSheet.getRow(fromCell.getRow());
		if (row == null)
			row = sourceSheet.createRow(fromCell.getRow());

		Cell cell = row.getCell(fromCell.getCol());
		if (cell == null)
			cell = row.createCell(fromCell.getCol());

		CreationHelper creationHelper = workbook.getCreationHelper();
		Hyperlink link = creationHelper.createHyperlink(HyperlinkType.DOCUMENT);

		// Address format: 'Sheet Name'!A1
		String cellAddress = "'" + targetSheetName + "'!" + getColumnLetter(targetCol) + (targetRow + 1);
		link.setAddress(cellAddress);

		cell.setHyperlink(link);
		cell.setCellValue(linkText);

		// Optional styling (blue underline)
		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setUnderline(Font.U_SINGLE);
		font.setColor(IndexedColors.BLUE.getIndex());
		style.setFont(font);
		cell.setCellStyle(style);
	}

	/**
	 * Creates a hyperlink from one cell to another within same or different sheet.
	 *
	 * @param source   The source cell reference.
	 * @param target   The destination cell reference.
	 * @param linkText The display text for the hyperlink.
	 *
	 *                 Example: CellRef from = new CellRef(sheet1, 0, 0); CellRef to
	 *                 = new CellRef(sheet2, 4, 2);
	 *                 helper.createHyperlinkToCell(from, to, "Go to Sheet2!C5");
	 */
	public void createHyperlinkToCell(CellRef source, CellRef target, String linkText) {
		createHyperlink(source.sheet, source.row, source.col, HyperlinkType.DOCUMENT, target.getAddress(), linkText);
	}

	/**
	 * Creates a hyperlink from the given cell to another cell in the same sheet.
	 *
	 * @param fromCell  Source cell reference (where hyperlink is placed)
	 * @param targetRow Target row index (0-based)
	 * @param targetCol Target column index (0-based)
	 * @param linkText  Display text for the hyperlink
	 */
	public void createHyperlinkToCellInSheet(CellRef fromCell, int targetRow, int targetCol, String linkText) {
		String address = fromCell.sheet.getSheetName() + "!$" + CellReference.convertNumToColString(targetCol) + "$"
				+ (targetRow + 1);
		createHyperlink(fromCell, HyperlinkType.DOCUMENT, address, linkText);
	}

	/**
	 * Creates a hyperlink from the given cell to a cell in another sheet.
	 *
	 * @param fromCell    Source cell reference (where hyperlink is placed)
	 * @param targetSheet Target sheet
	 * @param targetRow   Target row index
	 * @param targetCol   Target column index
	 * @param linkText    Display text for the hyperlink
	 */
	public void createHyperlinkToCellInAnotherSheet(CellRef fromCell, Sheet targetSheet, int targetRow, int targetCol,
			String linkText) {
		String sheetName = targetSheet.getSheetName();
		String address = "'" + sheetName + "'!$" + CellReference.convertNumToColString(targetCol) + "$"
				+ (targetRow + 1);
		createHyperlink(fromCell, HyperlinkType.DOCUMENT, address, linkText);
	}

	private void createHyperlink(CellRef cellRef, HyperlinkType type, String address, String text) {
		Cell cell = getOrCreateCell(cellRef.sheet, cellRef.row, cellRef.col);
		cell.setCellValue(text);

		CreationHelper creationHelper = workbook.getCreationHelper();
		Hyperlink hyperlink = creationHelper.createHyperlink(type);
		hyperlink.setAddress(address);
		cell.setHyperlink(hyperlink);

		CellStyle linkStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setUnderline(Font.U_SINGLE);
		font.setColor(IndexedColors.BLUE.getIndex());
		linkStyle.setFont(font);
		cell.setCellStyle(linkStyle);
	}

	Cell getOrCreateCell(Sheet sheet, int rowIdx, int colIdx) {
		Row row = sheet.getRow(rowIdx);
		if (row == null) {
			row = sheet.createRow(rowIdx);
		}
		Cell cell = row.getCell(colIdx);
		if (cell == null) {
			cell = row.createCell(colIdx);
		}
		return cell;
	}

	/**
	 * Reads a cell as a float. Returns null if the cell is empty, null, or not
	 * numeric.
	 * 
	 * @param sheet  The sheet containing the cell.
	 * @param rowNum The row index (0-based).
	 * @param colNum The column index (0-based).
	 * @return The float value, or null if the cell is not a number. Example Usage :
	 *         float myValue = readCellAsFloat(sheet, rowIndex, colIndex)
	 */
	public Float readCellAsFloat(Sheet sheet, int rowNum, int colNum) {

		Object value = readCell(sheet, rowNum, colNum);

		if (value instanceof Number) { // Checks if the cell value is number or not before conversion
			log(String.format("Reading Cell at row '%s' and column '%s' as float in sheet '%s'.", rowNum, colNum,
					sheet.getSheetName()), GREEN);
			return ((Number) value).floatValue();
		}

		if (value == null) { // null check
			log(String.format("Cell value is null at row '%s' and column '%s' in sheet '%s'.", rowNum, colNum,
					sheet.getSheetName()), YELLOW);

			return null;
		}

		log(String.format("Cannot read cell (%d, %d) as float. Value is not numeric: %s", rowNum, colNum,
				value.toString()), RED); // Error message
		return null;

	}

	public void formatCellAsPercent(Cell cell, int decimalPlaces) {
		CellStyle style = workbook.createCellStyle();
		style.setDataFormat(workbook.createDataFormat().getFormat("0." + "0".repeat(decimalPlaces) + "%"));
		cell.setCellStyle(style);
	}

	public void setBold(Cell cell) {
		Font boldFont = workbook.createFont();
		boldFont.setBold(true);

		CellStyle style = workbook.createCellStyle();
		style.cloneStyleFrom(cell.getCellStyle()); // Keep previous style if needed
		style.setFont(boldFont);
		cell.setCellStyle(style);
	}

}