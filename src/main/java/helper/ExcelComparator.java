package helper;



import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;
import sun.awt.image.ImageWatched;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

public class ExcelComparator {

    private static final String CELL_DATA_DOES_NOT_MATCH = "Cell Data does not Match ::";
    private static final String CELL_FONT_ATTRIBUTES_DOES_NOT_MATCH = "Cell Font Attributes does not Match ::";

    private static class Locator {
        Workbook workbook;
        Sheet sheet;
        Row row;
        Cell cell;
    }

    LinkedList<String> listOfDifferences = new LinkedList<>();
    List<String> stageFileNames = new ArrayList<>();
    List<String> prodFileNames = new ArrayList<>();


    public void getFileNameAndCompare(String stageFilePath, String prodFilePath) throws IOException, InvalidFormatException {
        getFileNameFromFolder();
        String result = "files//Summary//";
        for (int prodCount = 0; prodCount < prodFileNames.size(); prodCount++) {
            for (int stageCount = 0; stageCount < stageFileNames.size(); stageCount++) {
                if (prodFileNames.get(prodCount).contentEquals(stageFileNames.get(stageCount))) {
                    String outputFileName =  StringUtils.substringBefore(prodFileNames.get(prodCount), ".") +  "_" + System.currentTimeMillis() + ".xlsx";
                    String resultWorkBookPath = result + outputFileName;
                    FileInputStream Output_Excel;
                    FileInputStream PROD_Excel;
                    FileInputStream Stage_Excel;

                    File originalSample = new File(prodFilePath+prodFileNames.get(prodCount));
                    File output = new File(resultWorkBookPath);

                    Files.copy(originalSample.toPath(), output.toPath());

                    PROD_Excel = new FileInputStream(new File(prodFilePath+prodFileNames.get(prodCount)));
                    Stage_Excel = new FileInputStream(new File(stageFilePath+ stageFileNames.get(stageCount)));
                    Output_Excel = new FileInputStream(new File(resultWorkBookPath));

                    XSSFWorkbook stageWorkBook = new XSSFWorkbook(Stage_Excel);
                    XSSFWorkbook prodWorkBook = new XSSFWorkbook(PROD_Excel);
                    XSSFWorkbook outputWorkBook = new XSSFWorkbook(Output_Excel);

                    LinkedList<String> differences = ExcelComparator.compare(stageWorkBook, prodWorkBook,outputWorkBook);
                    createSummarySheet(outputWorkBook, differences, resultWorkBookPath);
                    if (differences.isEmpty()) {
                        System.out.println("Both files are identical");
                    } else {
                        System.out.println("---------------------------------------------------------------------");
                        System.out.println("Comparison summary for " + prodFileNames.get(prodCount));
                        System.out.println("Total differences found :" + differences.size());
                        System.out.println("---------------------------------------------------------------------");
                        for(String s: differences) {
                            System.out.println(s);
                            System.out.println("---------------------------------------------------------------------");
                        }
                    }
               }
            }
        }
    }

    private static void createSummarySheet(XSSFWorkbook workbook3, List<String> differences, String resultPath) throws IOException {
        XSSFSheet sheet = workbook3.createSheet("Summary");
        XSSFRow row;
        LinkedHashMap<String, Object[]> excelCompareData  = new LinkedHashMap<>();
        excelCompareData.put("1", new Object[] { "Summary" });
        for(int i =0; i < differences.size(); i++) {
            excelCompareData.put(String.valueOf(i), new Object[] { differences.get(i) });
        }
        int rowid = 0;
        Set<String> keyid = excelCompareData.keySet();
        for (String key : keyid) {
            row = sheet.createRow(rowid++);
            Object[] objectArr = excelCompareData.get(key);
            int cellid = 0;
            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String)obj);
            }
        }
        FileOutputStream output_file = new FileOutputStream(new File(resultPath));
        workbook3.write(output_file);
        output_file.close();
    }

    public void getFileNameFromFolder() {
        String stageFilePath = "files//Stage//";
        String prodFilePath = "files//Prod//";

        File stageFolder = new File(stageFilePath);
        File[] listOfStageFiles = stageFolder.listFiles();
//        System.out.println("Files present inside Stage folder : ");
        for (int stageCount = 0; stageCount < listOfStageFiles.length; stageCount++) {
            if (listOfStageFiles[stageCount].isFile()) {
//                System.out.println(stageCount + " : " + listOfStageFiles[stageCount].getName());
                stageFileNames.add(listOfStageFiles[stageCount].getName());
            }
        }
        File prodFolder = new File(prodFilePath);
        File[] listOfProdFiles = prodFolder.listFiles();
//        System.out.println("Files present inside Prod folder : ");
        for (int prodCount = 0; prodCount < listOfProdFiles.length; prodCount++) {
            if (listOfProdFiles[prodCount].isFile()) {
//                System.out.println(prodCount + " : " + listOfProdFiles[prodCount].getName());
                prodFileNames.add(listOfProdFiles[prodCount].getName());
            }
        }
    }

    /**
     * Utility to compare Excel File Contents cell by cell for all sheets.
     *
     * @param wb1 the workbook1
     * @param wb2 the workbook2
     * @return the Excel file difference containing a flag and a list of differences
     */
    public static LinkedList<String> compare(Workbook wb1, Workbook wb2, Workbook wb3) throws IOException {
        Locator loc1 = new Locator();
        Locator loc2 = new Locator();
        Locator loc3 = new Locator();
        loc1.workbook = wb1;
        loc2.workbook = wb2;
        loc3.workbook = wb3;

        ExcelComparator excelComparator = new ExcelComparator();

        excelComparator.compareNumberOfSheets(loc1, loc2 );
        excelComparator.compareSheetNames(loc1, loc2);
        excelComparator.compareSheetData(loc1, loc2, loc3);

        return excelComparator.listOfDifferences;
    }

    /**
     * Compare data in all sheets.
     */
    private void compareDataInAllSheets(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        for (int i = 0; i < loc1.workbook.getNumberOfSheets(); i++) {
            if (loc2.workbook.getNumberOfSheets() <= i) return;

            loc1.sheet = loc1.workbook.getSheetAt(i);
            loc2.sheet = loc2.workbook.getSheetAt(i);
            loc3.sheet = loc3.workbook.getSheetAt(i);
            compareDataInSheet(loc1, loc2, loc3);
        }
    }

    private void compareDataInSheet(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        for (int j = 0; j < loc1.sheet.getPhysicalNumberOfRows(); j++) {
            if (loc2.sheet.getPhysicalNumberOfRows() <= j) return;

            loc1.row = loc1.sheet.getRow(j);
            loc2.row = loc2.sheet.getRow(j);
            loc3.row = loc3.sheet.getRow(j);

            if ((loc1.row == null) || (loc2.row == null)) {
                continue;
            }

            compareDataInRow(loc1, loc2, loc3);
        }
    }

    private void compareDataInRow(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        for (int k = 0; k < loc1.row.getLastCellNum(); k++) {
            if (loc2.row.getPhysicalNumberOfCells() <= k) return;

            loc1.cell = loc1.row.getCell(k);
            loc2.cell = loc2.row.getCell(k);
            loc3.cell = loc3.row.getCell(k);

            if ((loc1.cell == null) || (loc2.cell == null)) {
                continue;
            }
            compareDataInCell(loc1, loc2, loc3);
        }
    }

    private void compareDataInCell(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        if (isCellTypeMatches(loc1, loc2, loc3)) {
            final CellType loc1cellType = loc1.cell.getCellType();
            switch(loc1cellType) {
                case BLANK:
                case STRING:
                case ERROR:
                    isCellContentMatches(loc1,loc2, loc3);
                    break;
                case BOOLEAN:
                    isCellContentMatchesForBoolean(loc1,loc2, loc3);
                    break;
                case FORMULA:
                    isCellContentMatchesForFormula(loc1,loc2, loc3);
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(loc1.cell)) {
                        isCellContentMatchesForDate(loc1,loc2, loc3);
                    } else {
                        isCellContentMatchesForNumeric(loc1,loc2, loc3);
                    }
                    break;
                default:
                    throw new IllegalStateException("Unexpected cell type: " + loc1cellType);
            }
        }
        isCellFillPatternMatches(loc1,loc2, loc3);
        isCellAlignmentMatches(loc1,loc2, loc3);
        isCellHiddenMatches(loc1,loc2, loc3);
        isCellLockedMatches(loc1,loc2, loc3);
        isCellFontFamilyMatches(loc1,loc2, loc3);
        isCellFontSizeMatches(loc1,loc2, loc3);
        isCellFontBoldMatches(loc1,loc2, loc3);
        isCellUnderLineMatches(loc1,loc2, loc3);
        isCellFontItalicsMatches(loc1,loc2, loc3);
        isCellBorderMatches(loc1,loc2, loc3,'t');
        isCellBorderMatches(loc1,loc2,loc3,'l');
        isCellBorderMatches(loc1,loc2, loc3,'b');
        isCellBorderMatches(loc1,loc2, loc3,'r');
        isCellFillBackGroundMatches(loc1,loc2, loc3);
    }

    /**
     * Compare number of columns in sheets.
     */
    private void compareNumberOfColumnsInSheets(Locator loc1, Locator loc2) {
        for (int i = 0; i < loc1.workbook.getNumberOfSheets(); i++) {
            if (loc2.workbook.getNumberOfSheets() <= i) return;

            loc1.sheet = loc1.workbook.getSheetAt(i);
            loc2.sheet = loc2.workbook.getSheetAt(i);

            Iterator<Row> ri1 = loc1.sheet.rowIterator();
            Iterator<Row> ri2 = loc2.sheet.rowIterator();

            int num1 = (ri1.hasNext()) ? ri1.next().getPhysicalNumberOfCells() : 0;
            int num2 = (ri2.hasNext()) ? ri2.next().getPhysicalNumberOfCells() : 0;

            if (num1 != num2) {
                String str = String.format(Locale.ROOT, "%s\nworkbook1 -> %s [%d] != workbook2 -> %s [%d]",
                        "Number Of Columns does not Match ::",
                        loc1.sheet.getSheetName(), num1,
                        loc2.sheet.getSheetName(), num2
                );
                listOfDifferences.add(str);
            }
        }
    }

    /**
     * Compare number of rows in sheets.
     */
    private void compareNumberOfRowsInSheets(Locator loc1, Locator loc2) {
        for (int i = 0; i < loc1.workbook.getNumberOfSheets(); i++) {
            if (loc2.workbook.getNumberOfSheets() <= i) return;

            loc1.sheet = loc1.workbook.getSheetAt(i);
            loc2.sheet = loc2.workbook.getSheetAt(i);

            int num1 = loc1.sheet.getPhysicalNumberOfRows();
            int num2 = loc2.sheet.getPhysicalNumberOfRows();

            if (num1 != num2) {
                String str = String.format(Locale.ROOT, "%s\nworkbook1 -> %s [%d] != workbook2 -> %s [%d]",
                        "Number Of Rows does not Match ::",
                        loc1.sheet.getSheetName(), num1,
                        loc2.sheet.getSheetName(), num2
                );
                listOfDifferences.add(str);
            }
        }

    }

    /**
     * Compare number of sheets.
     */
    private void compareNumberOfSheets(Locator loc1, Locator loc2) {
        int num1 = loc1.workbook.getNumberOfSheets();
        int num2 = loc2.workbook.getNumberOfSheets();
        if (num1 != num2) {
            String str = String.format(Locale.ROOT, "%s\nworkbook1 [%d] != workbook2 [%d]",
                    "Number of Sheets do not match ::",
                    num1, num2
            );
            listOfDifferences.add(str);
        }
    }

    /**
     * Compare sheet data.
     */
    private void compareSheetData(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        compareNumberOfRowsInSheets(loc1, loc2);
        compareNumberOfColumnsInSheets(loc1, loc2);
        compareDataInAllSheets(loc1, loc2, loc3);

    }

    /**
     * Compare sheet names.
     */
    private void compareSheetNames(Locator loc1, Locator loc2) {
        for (int i = 0; i < loc1.workbook.getNumberOfSheets(); i++) {
            String name1 = loc1.workbook.getSheetName(i);
            String name2 = (loc2.workbook.getNumberOfSheets() > i) ? loc2.workbook.getSheetName(i) : "";

            if (!name1.equals(name2)) {
                String str = String.format(Locale.ROOT, "%s\nworkbook1 -> %s [%d] != workbook2 -> %s [%d]",
                        "Name of the sheets do not match ::", name1, i+1, name2, i+1
                );
                listOfDifferences.add(str);
            }
        }
    }

    public void addCommentInCell(Locator loc3, String value, String sheetName) throws IOException {
        Workbook workbook1 = loc3.workbook;
        Sheet sheet = loc3.workbook.getSheet(sheetName);
        CreationHelper factory = workbook1.getCreationHelper();
        ClientAnchor anchor = factory.createClientAnchor();
        Drawing drawing = sheet.createDrawingPatriarch();

        Cell cell = loc3.cell;
        Row row  = loc3.row;

        anchor.setCol1(cell.getColumnIndex()+1);                           // starts at column A + 1 = B
        anchor.setDx1(10* Units.EMU_PER_PIXEL); // plus 10 px
        anchor.setCol2(cell.getColumnIndex()+2);                           // ends at column A + 2 = C
        anchor.setDx2(10*Units.EMU_PER_PIXEL); // plus 10 px
        anchor.setRow1(row.getRowNum());                                   // starts at row 1
        anchor.setDy1(10*Units.EMU_PER_PIXEL); // plus 10 px
        anchor.setRow2(row.getRowNum()+3);                                 // ends at row 4
        anchor.setDy2(10*Units.EMU_PER_PIXEL); // plus 10 px

        Comment comment = drawing.createCellComment(anchor);
        RichTextString str = factory.createRichTextString(value);
        comment.setString(str);
        cell.setCellComment(comment);
        saveSheet(loc3);
    }

    public void addSummary(Locator loc3, String value, String sheetName) throws IOException {
        Workbook workbook1 = loc3.workbook;
        Sheet sheet = loc3.workbook.getSheet(sheetName);
        CreationHelper factory = workbook1.getCreationHelper();
        ClientAnchor anchor = factory.createClientAnchor();
        Drawing drawing = sheet.createDrawingPatriarch();
        Cell cell = sheet.getRow(0).getCell(0);
        Row row = sheet.getRow(0);
        anchor.setCol1(cell.getColumnIndex()+6);                           // starts at column A + 1 = B
        anchor.setDx1(10* Units.EMU_PER_PIXEL); // plus 10 px
        anchor.setCol2(cell.getColumnIndex()+4);                           // ends at column A + 2 = C
        anchor.setDx2(10*Units.EMU_PER_PIXEL); // plus 10 px
        anchor.setRow1(row.getRowNum());                                   // starts at row 1
        anchor.setDy1(10*Units.EMU_PER_PIXEL); // plus 10 px
        anchor.setRow2(row.getRowNum()+6);                                 // ends at row 4
        anchor.setDy2(10*Units.EMU_PER_PIXEL); // plus 10 px

        Comment comment = drawing.createCellComment(anchor);
        RichTextString str = factory.createRichTextString(value);
        comment.setString(str);
        cell.setCellComment(comment);
        saveSheet(loc3);
    }

    public void saveSheet(Locator loc3) throws IOException {
        Workbook workbook1 = loc3.workbook;
        FileOutputStream output_file = new FileOutputStream(new File("src//main//resources//output.xlsx"));
        workbook1.write(output_file);
        output_file.close();
    }

    public void changeBackGroundColor(Locator loc3) throws IOException {
        Workbook workbook1 = loc3.workbook;
        XSSFCellStyle tCs = (XSSFCellStyle) workbook1.createCellStyle();
        tCs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        tCs.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        loc3.cell.setCellStyle(tCs);
        saveSheet(loc3);
    }

    /**
     * Formats the message.
     */
    private void addMessage(Locator loc1, Locator loc2,Locator loc3, String messageStart, String value1, String value2) throws IOException {
        String str =
                String.format(Locale.ROOT, "%s\nworkbook1 -> %s -> %s [%s] != workbook2 -> %s -> %s [%s]",
                        messageStart,
                        loc1.sheet.getSheetName(), new CellReference(loc1.cell).formatAsString(), value1,
                        loc2.sheet.getSheetName(), new CellReference(loc2.cell).formatAsString(), value2
                );
        listOfDifferences.add(str);
        changeBackGroundColor(loc3);
        addCommentInCell(loc3,value1, loc3.sheet.getSheetName());

    }

    /**
     * Checks if cell alignment matches.
     */
    private void isCellAlignmentMatches(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        // TODO: check for NPE
        HorizontalAlignment align1 = loc1.cell.getCellStyle().getAlignment();
        HorizontalAlignment align2 = loc2.cell.getCellStyle().getAlignment();
        if (align1 != align2) {
            addMessage(loc1, loc2, loc3,
                    "Cell Alignment does not Match ::",
                    align1.name(),
                    align2.name()
            );
        }
    }

    /**
     * Checks if cell border bottom matches.
     */
    private void isCellBorderMatches(Locator loc1, Locator loc2, Locator loc3, char borderSide) throws IOException {
        if (!(loc1.cell instanceof XSSFCell)) return;
        XSSFCellStyle style1 = ((XSSFCell)loc1.cell).getCellStyle();
        XSSFCellStyle style2 = ((XSSFCell)loc2.cell).getCellStyle();
        boolean b1, b2;
        String borderName;
        switch (borderSide) {
            case 't': default:
                b1 = style1.getBorderTop() == BorderStyle.THIN;
                b2 = style2.getBorderTop() == BorderStyle.THIN;
                borderName = "TOP";
                break;
            case 'b':
                b1 = style1.getBorderBottom() == BorderStyle.THIN;
                b2 = style2.getBorderBottom() == BorderStyle.THIN;
                borderName = "BOTTOM";
                break;
            case 'l':
                b1 = style1.getBorderLeft() == BorderStyle.THIN;
                b2 = style2.getBorderLeft() == BorderStyle.THIN;
                borderName = "LEFT";
                break;
            case 'r':
                b1 = style1.getBorderRight() == BorderStyle.THIN;
                b2 = style2.getBorderRight() == BorderStyle.THIN;
                borderName = "RIGHT";
                break;
        }
        if (b1 != b2) {
            addMessage(loc1, loc2, loc3,
                    "Cell Border Attributes does not Match ::",
                    (b1 ? "" : "NOT ")+borderName+" BORDER",
                    (b2 ? "" : "NOT ")+borderName+" BORDER"
            );
        }
    }

    /**
     * Checks if cell content matches.
     */
    private void isCellContentMatches(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        // TODO: check for null and non-rich-text cells
        String str1 = loc1.cell.getRichStringCellValue().getString();
        String str2 = loc2.cell.getRichStringCellValue().getString();
        if (!str1.equals(str2)) {
            addMessage(loc1,loc2,loc3,CELL_DATA_DOES_NOT_MATCH,str1,str2);
        }
    }

    /**
     * Checks if cell content matches for boolean.
     */
    private void isCellContentMatchesForBoolean(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        boolean b1 = loc1.cell.getBooleanCellValue();
        boolean b2 = loc2.cell.getBooleanCellValue();
        if (b1 != b2) {
            addMessage(loc1,loc2, loc3, CELL_DATA_DOES_NOT_MATCH,Boolean.toString(b1),Boolean.toString(b2));
        }
    }

    /**
     * Checks if cell content matches for date.
     */
    private void isCellContentMatchesForDate(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        Date date1 = loc1.cell.getDateCellValue();
        Date date2 = loc2.cell.getDateCellValue();
        if (!date1.equals(date2)) {
            addMessage(loc1, loc2, loc3, CELL_DATA_DOES_NOT_MATCH, date1.toString(), date2.toString());
        }
    }


    /**
     * Checks if cell content matches for formula.
     */
    private void isCellContentMatchesForFormula(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        // TODO: actually evaluate the formula / NPE checks
        String form1 = loc1.cell.getCellFormula();
        String form2 = loc2.cell.getCellFormula();
        if (!form1.equals(form2)) {
            addMessage(loc1, loc2,loc3, CELL_DATA_DOES_NOT_MATCH, form1, form2);
        }
    }

    /**
     * Checks if cell content matches for numeric.
     */
    private void isCellContentMatchesForNumeric(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        // TODO: Check for NaN
        double num1 = loc1.cell.getNumericCellValue();
        double num2 = loc2.cell.getNumericCellValue();
        if (num1 != num2) {
            addMessage(loc1, loc2,loc3, CELL_DATA_DOES_NOT_MATCH, Double.toString(num1), Double.toString(num2));
        }
    }

    private String getCellFillBackground(Locator loc) {
        Color col = loc.cell.getCellStyle().getFillForegroundColorColor();
        return (col instanceof XSSFColor) ? ((XSSFColor)col).getARGBHex() : "NO COLOR";
    }

    /**
     * Checks if cell file back ground matches.
     */
    private void isCellFillBackGroundMatches(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        String col1 = getCellFillBackground(loc1);
        String col2 = getCellFillBackground(loc2);
        if (!col1.equals(col2)) {
            addMessage(loc1, loc2, loc3, "Cell Fill Color does not Match ::", col1, col2);
        }
    }
    /**
     * Checks if cell fill pattern matches.
     */
    private void isCellFillPatternMatches(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        // TOOO: Check for NPE
        FillPatternType fill1 = loc1.cell.getCellStyle().getFillPattern();
        FillPatternType fill2 = loc2.cell.getCellStyle().getFillPattern();
        if (fill1 != fill2) {
            addMessage(loc1, loc2,loc3,
                    "Cell Fill pattern does not Match ::",
                    fill1.name(),
                    fill2.name()
            );
        }
    }

    /**
     * Checks if cell font bold matches.
     */
    private void isCellFontBoldMatches(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        if (!(loc1.cell instanceof XSSFCell)) return;
        boolean b1 = ((XSSFCell)loc1.cell).getCellStyle().getFont().getBold();
        boolean b2 = ((XSSFCell)loc2.cell).getCellStyle().getFont().getBold();
        if (b1 != b2) {
            addMessage(loc1, loc2, loc3,
                    CELL_FONT_ATTRIBUTES_DOES_NOT_MATCH,
                    (b1 ? "" : "NOT ")+"BOLD",
                    (b2 ? "" : "NOT ")+"BOLD"
            );
        }
    }

    /**
     * Checks if cell font family matches.
     */
    private void isCellFontFamilyMatches(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        // TODO: Check for NPEs
        if (!(loc1.cell instanceof XSSFCell)) return;
        String family1 = ((XSSFCell)loc1.cell).getCellStyle().getFont().getFontName();
        String family2 = ((XSSFCell)loc2.cell).getCellStyle().getFont().getFontName();
        if (!family1.equals(family2)) {
            addMessage(loc1, loc2, loc3,"Cell Font Family does not Match ::", family1, family2);
        }
    }

    /**
     * Checks if cell font italics matches.
     */
    private void isCellFontItalicsMatches(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        if (!(loc1.cell instanceof XSSFCell)) return;
        boolean b1 = ((XSSFCell)loc1.cell).getCellStyle().getFont().getItalic();
        boolean b2 = ((XSSFCell)loc2.cell).getCellStyle().getFont().getItalic();
        if (b1 != b2) {
            addMessage(loc1, loc2, loc3,
                    CELL_FONT_ATTRIBUTES_DOES_NOT_MATCH,
                    (b1 ? "" : "NOT ")+"ITALICS",
                    (b2 ? "" : "NOT ")+"ITALICS"
            );
        }
    }

    /**
     * Checks if cell font size matches.
     */
    private void isCellFontSizeMatches(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        if (!(loc1.cell instanceof XSSFCell)) return;
        short size1 = ((XSSFCell)loc1.cell).getCellStyle().getFont().getFontHeightInPoints();
        short size2 = ((XSSFCell)loc2.cell).getCellStyle().getFont().getFontHeightInPoints();
        if (size1 != size2) {
            addMessage(loc1, loc2, loc3,
                    "Cell Font Size does not Match ::",
                    Short.toString(size1),
                    Short.toString(size2)
            );
        }
    }

    /**
     * Checks if cell hidden matches.
     */
    private void isCellHiddenMatches(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        boolean b1 = loc1.cell.getCellStyle().getHidden();
        boolean b2 = loc1.cell.getCellStyle().getHidden();
        if (b1 != b2) {
            addMessage(loc1, loc2, loc3,
                    "Cell Visibility does not Match ::",
                    (b1 ? "" : "NOT ")+"HIDDEN",
                    (b2 ? "" : "NOT ")+"HIDDEN"
            );
        }
    }

    /**
     * Checks if cell locked matches.
     */
    private void isCellLockedMatches(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        boolean b1 = loc1.cell.getCellStyle().getLocked();
        boolean b2 = loc1.cell.getCellStyle().getLocked();
        if (b1 != b2) {
            addMessage(loc1, loc2, loc3,
                    "Cell Protection does not Match ::",
                    (b1 ? "" : "NOT ")+"LOCKED",
                    (b2 ? "" : "NOT ")+"LOCKED"
            );
        }
    }

    /**
     * Checks if cell type matches.
     */
    private boolean isCellTypeMatches(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        CellType type1 = loc1.cell.getCellType();
        CellType type2 = loc2.cell.getCellType();
        if (type1 == type2) return true;
        addMessage(loc1, loc2, loc3,
                "Cell Data-Type does not Match in :: ",
                type1.name(), type2.name()
        );
        return false;
    }

    /**
     * Checks if cell under line matches.
     */
    private void isCellUnderLineMatches(Locator loc1, Locator loc2, Locator loc3) throws IOException {
        // TOOO: distinguish underline type
        if (!(loc1.cell instanceof XSSFCell)) return;
        byte b1 = ((XSSFCell)loc1.cell).getCellStyle().getFont().getUnderline();
        byte b2 = ((XSSFCell)loc2.cell).getCellStyle().getFont().getUnderline();
        if (b1 != b2) {
            addMessage(loc1, loc2, loc3,
                    CELL_FONT_ATTRIBUTES_DOES_NOT_MATCH,
                    (b1 == 1 ? "" : "NOT ")+"UNDERLINE",
                    (b2 == 1 ? "" : "NOT ")+"UNDERLINE"
            );
        }
    }
}
