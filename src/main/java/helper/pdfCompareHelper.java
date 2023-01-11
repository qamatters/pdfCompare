package helper;

import util.PDFUtilHelper;

import java.io.IOException;
import java.util.List;

public class pdfCompareHelper {

    static PDFUtilHelper pdfutil = new PDFUtilHelper();

    public static String getPageTextFromPDF(String report, int startPageNum, int endPageNum) throws IOException {
        return  pdfutil.getText(report,startPageNum,endPageNum);
    }

    public static List<String> comparePDFImageSpecificPage(String file, int startPage, int endPage, String folderName) throws IOException {
        return pdfutil.savePdfAsImage(file,startPage,endPage,folderName);

    }

    public static String getSpecificTextFromPDF(String report, int startPageNum, int endPageNum) throws IOException {
        return pdfutil.getPdfPageText(report,startPageNum,endPageNum);

    }

}
