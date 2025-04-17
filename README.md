# 📊 ExcelHelper Utility for Java (Apache POI)

A robust and feature-rich utility for creating, reading, and managing Excel files (`.xls`, `.xlsx`) using Apache POI. It is designed to simplify working with Excel files — whether you're generating styled reports, importing/exporting data, or building automation test dashboards.

---

## 🚀 Features

- 📄 **Read/Write** Excel files in `.xls`, `.xlsx` (HSSF, XSSF, SXSSF)
- 📊 **Create Reports**: Test reports, sales data, database exports, dashboards
- 🎨 **Rich Styling**: Headers, fonts, borders, background colors, formatting
- 🧪 **Automation Reports**: With PASS/FAIL status coloring & hyperlinks
- 🔗 **Hyperlinks** to cells/sheets
- 🔃 **Import/Export** CSV
- 🧱 **Merge Cells**, Freeze Panes, Auto-Filters
- 📈 **Handle Large Data Sets** with streaming write (`SXSSFWorkbook`)
- 🛡 **Protect Sheets**, Lock Cells
- 🔍 **Cell Utilities**: Validation, emptiness check, column name/index conversion

---

## 📁 Project Structure

```bash
Excel-Utilities/
├── src/
│   └── main/java/com/excelUtilities/ExcelHelper.java     # The main helper class
│   └── test/java/com/excelUtilities/ExcelHelperTest.java # JUnit test suite
├── resources/
├── README.md
├── pom.xml or build.gradle (if using Maven/Gradle)


🧑‍💻 Getting Started

✅ Prerequisites

Java 11+
Maven or Gradle
Apache POI libraries:
poi
poi-ooxml
poi-ooxml-schemas
poi-scratchpad (if using charts or rich text)
commons-collections4
xmlbeans

🧩 Maven Dependency (if needed)

<dependencies>
  <dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.2.3</version>
  </dependency>
</dependencies>

🛠 Example Usage

ExcelHelper helper = new ExcelHelper();
helper.setFilePath("Report.xlsx", ExcelMode.WRITE);

Sheet sheet = helper.createSheet("Summary");
helper.writeHeaders(sheet, List.of("ID", "Name", "Status"));
helper.writeRow(sheet, 1, List.of("TC-001", "Login Test", "PASS"));

helper.saveWorkbook("Report.xlsx");
helper.closeWorkbook();

🧪 Running Tests

# If using Maven
mvn test

# If using Gradle
gradle test


The ExcelHelperTest.java contains over 40 unit tests including:

Cell writing/reading (all types)

Styling (headers, borders, colors)

Large data handling

Automation and business report generation

Hyperlinks and sheet navigation

📦 Sample Reports Generated

✅ Automation Test Reports (with links to test cases)

📈 Sales Reports (monthly budgets, actuals, %)

🗃 Database Table Exports

📊 Monitoring Dashboards

🧾 Comparative Tax Summary Reports (styled like official formats)

📌 Notable Utilities
setBold(Cell)

formatCellAsCurrency, formatCellAsDate, formatCellAsPercent

autoSizeAllColumns(Sheet)

createHyperlinkToCellInDifferentSheet(...)

writeLargeData(...) — optimized for millions of rows

🤝 Contribution
Contributions are welcome! To contribute:

Fork the repo

Create your feature branch: git checkout -b feature/my-feature

Commit your changes: git commit -m 'Add some feature'

Push to the branch: git push origin feature/my-feature

Open a Pull Request

📄 License
MIT License © 2025
Maintained by [Mr.DK]

💬 Questions?
For help or suggestions, open an issue or contact me on [LinkedIn/GitHub/email].


Would you like to:

- Auto-generate `.gitignore` (e.g., exclude `.idea/`, `.classpath`, `target/`, etc.)?
- Add a sample screenshot of a report?
- Include badge links (build passing, license, etc.)?

Let me know and I’ll plug them right in!