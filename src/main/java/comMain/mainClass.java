package comMain;

import helper.ExcelComparator;
import helper.commonHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.util.LinkedList;

import static helper.commonHelper.getFileNameFromFolder;

public class mainClass {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        String stageFilePath = "files//Stage//";
        String prodFilePath = "files//Prod//";
        commonHelper helper = new commonHelper();
        ExcelComparator excelHelper = new ExcelComparator();

        LinkedList<String> stageFileNames = getFileNameFromFolder(stageFilePath);
        LinkedList<String> prodFileNames = getFileNameFromFolder(prodFilePath);
        if (stageFileNames.stream().allMatch(file -> file.endsWith(".pdf")) && prodFileNames.stream().allMatch(file -> file.endsWith(".pdf"))) {
            helper.getFileNameAndCompare();
        } else if (stageFileNames.stream().allMatch(file -> file.endsWith(".xlsx")) && prodFileNames.stream().allMatch(file -> file.endsWith(".xlsx"))) {
            excelHelper.getFileNameAndCompare(stageFilePath, prodFilePath);
        } else {
            System.out.println("Please place either pdf or excel files inside folders for comparison. Don't keep both type files");
        }
    }
}
