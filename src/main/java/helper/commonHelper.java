package helper;

import model.PageModel;
import net.sourceforge.tess4j.Tesseract;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.comparator.LastModifiedFileComparator;
import org.apache.commons.lang3.StringUtils;
import org.apache.pdfbox.cos.COSDictionary;
import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.multipdf.LayerUtility;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.form.PDFormXObject;
import util.PDFUtilHelper;

import java.awt.geom.AffineTransform;
import java.io.*;
import java.util.*;

import static helper.imageCompareHelper.compareWithBaseImage;
import static helper.imageCompareHelper.createPDFFromSingleFile;

public class commonHelper {

    List<String> stageFileNames = new ArrayList<>();
    List<String> prodFileNames = new ArrayList<>();
    List<String> diffOutputFileNames = new ArrayList<>();
    String productionFileName = "";


    public void getFileNameAndCompare() throws IOException {
        getFileNameFromFolder();
        String finalReportPath = "files//Report//";
        String result = "files//Summary//";
        for (int prodCount = 0; prodCount < prodFileNames.size(); prodCount++) {
            for (int stageCount = 0; stageCount < stageFileNames.size(); stageCount++) {
                if (prodFileNames.get(prodCount).contentEquals(stageFileNames.get(stageCount))) {
                    checkForFileContent(prodFileNames.get(prodCount), stageFileNames.get(stageCount));
                    mergePdf(finalReportPath, result, productionFileName);
                    deleteTempFiles();
                    break;
                }
            }
        }
    }

    private void mergePdf(String filePath, String finalReportPath, String reportName) throws IOException {
        File folder = new File(filePath);
        File[] listOdPDFFiles = folder.listFiles();
        Arrays.sort(listOdPDFFiles, LastModifiedFileComparator.LASTMODIFIED_COMPARATOR);
        PDFMergerUtility pdfMergerUtility = new PDFMergerUtility();
        pdfMergerUtility.setDestinationFileName(finalReportPath + reportName);
        for (File file : listOdPDFFiles) {
            pdfMergerUtility.addSource(file);
        }
        pdfMergerUtility.mergeDocuments(MemoryUsageSetting.setupTempFileOnly());
//        System.out.println("PDF Documents merged to a single file");
    }

    private void checkForFileContent(String prodFile, String stageFile) throws IOException {
        String stagePath = "files//Stage//";
        String prodPath = "files//Prod//";
        diffOutputFileNames = new ArrayList<>();

        List<PageModel> prodPages = new ArrayList<>();
        List<PageModel> stagePages = new ArrayList<>();

        productionFileName = prodFile;
        System.out.println("-----------------------------------------");
        System.out.println(prodFile + " file comparison started");
        System.out.println("-----------------------------------------");
        String[] titleArray = getIgnoredTitles().split(",");

        System.out.println("Titles to ignore while comparison");
        if (!titleArray[0].equals("")) {
            for (String s : titleArray) {
                System.out.println(s + "\n");
            }
        }

        String stageReport = stagePath + stageFile;
        String prodReport = prodPath + prodFile;


        ignorePagesByTitle(prodReport, stageReport, titleArray, prodPages, stagePages);


        for (int i = 0, j = 0; i < prodPages.size() && j < stagePages.size(); i++, j++) {
            String prodPageText = prodPages.get(i).getPageText();
            String stageText = stagePages.get(j).getPageText();

            String[] prodpgaeWords = prodPageText.split("\\s");
            String[] stagegaeWords = stageText.split("\\s");


            if(prodpgaeWords.length< stagegaeWords.length){
                HashSet<String> s1 = new HashSet<String>(Arrays.asList(stagegaeWords));
                s1.removeAll(Arrays.asList(prodpgaeWords));
                System.out.println(s1);
            } else {
                HashSet<String> s1 = new HashSet<String>(Arrays.asList(prodpgaeWords));
                s1.removeAll(Arrays.asList(stagegaeWords));
                System.out.println(s1);
            }



//            while (i < 2) {
//                comparePages(prodPageText, stageText, prodReport, stageReport, prodPages.get(i).getPageNumber(),
//                        stagePages.get(i).getPageNumber());
//                j = j + 1;
//                i++;
//                prodPageText = prodPages.get(i).getPageText();
//                stageText = stagePages.get(j).getPageText();
//            }
//
//            if (i > 2) {
                String prodPageTitle = prodPages.get(i).getPageTitle();
                String stagePageTitle = stagePages.get(j).getPageTitle();
                // If Page title is same then compare those pages and increment page number of both prod and stage by 1
                if (prodPageTitle.equals(stagePageTitle)) {
                    comparePages(prodPageText, stageText, prodReport, stageReport, prodPages.get(i).getPageNumber(),
                            stagePages.get(j).getPageNumber());
                } else {
                    System.out.println(prodPageTitle.trim() + " title page from production report (Page " + i + ") is not present in stage report (Page " + j + ")");
                    if (j < stagePages.size() - 1) {
                        j = j + 1;
                        stagePageTitle = stagePages.get(j).getPageTitle();
                        if (prodPageTitle.equals(stagePageTitle)) {
                            comparePages(prodPageText, stageText, prodReport, stageReport, prodPages.get(i).getPageNumber(),
                                    stagePages.get(j).getPageNumber());
                        } else {
                            j = j - 1;
                            System.out.println(prodPageTitle.trim() + " title page from production report (Page " + i + ") is not present in stage report (Page " + j + ")");
                        }
                    }

                }

//            }
        }
        System.out.println("-----------------------------------------");
        System.out.println(prodFile + " File comparison ended");
        System.out.println("-----------------------------------------");

    }

    private void comparePages(String prodPageText, String stageText, String prodReport, String stageReport, int i, int j) {
        try {
            List<String> pageAsImageFromProd = pdfCompareHelper.comparePDFImageSpecificPage(prodReport, i, i, "prod");
            List<String> pageAsImageFromStage = pdfCompareHelper.comparePDFImageSpecificPage(stageReport, j, j, "stage");

            String diffFileName = "diff" + "_" + i + "_" + System.currentTimeMillis();
            String prodTemp = "prodTemp" + "_" + i + "_" + System.currentTimeMillis() + ".pdf";
            String stageTemp = "stageTemp" + "_" + i + "_" + System.currentTimeMillis() + ".pdf";
            String diffTemp = "diff_File" + "_" + i + "_" + System.currentTimeMillis() + ".pdf";
            String firstmerge = "firstmerged_temp" + "_" + i + "_" + System.currentTimeMillis() + ".pdf";
            String diffOutPut = compareWithBaseImage(new File(pageAsImageFromStage.get(0)), new File(pageAsImageFromProd.get(0)), diffFileName );

            if (diffOutPut.equals("N")) { // If pages are same and there are no differences
                System.out.println("Page " + i + " from Production content matches with stage page " + j);
            } else {
                System.out.println("Page " + i + " from Production content does not match with stage page " + j);
                createPDFFromSingleFile(pageAsImageFromProd.get(0), prodTemp);
                createPDFFromSingleFile(pageAsImageFromStage.get(0), stageTemp);
                createPDFFromSingleFile(diffOutPut, diffTemp);

                String finalResultPath = "files//Results//";
                String finalPathFromMergedFile = finalResultPath + firstmerge;
                String finalReportPath = "files//Report//";
                String finalMergedFileName = "finalTemp" + "_" + i + "_" + System.currentTimeMillis() + ".pdf";
                mergedFileSideBySide(prodTemp, stageTemp, finalResultPath, firstmerge, "left", "right");
                mergedFileSideBySide(finalPathFromMergedFile, diffTemp, finalReportPath, finalMergedFileName, "top", "bottom");
            }

        } catch (AssertionError | IOException e) {
            System.out.println("-------------------------");
            System.out.println(e.getMessage());
            System.out.println("--------------------------");
        }
        System.out.println("-----------------------------");

    }

    private void mergedFileSideBySide(String file1Path, String file2Path, String mergedFilePath, String mergedFileName, String layerName1, String layerName2)
            throws IOException {

        File pdf1File = new File(file1Path);
        File pdf2File = new File(file2Path);
        File outPdfFile = new File(mergedFilePath);

        PDDocument pdf1 = null;
        PDDocument pdf2 = null;
        PDDocument outPdf = null;

        try {
            pdf1 = PDDocument.load(pdf1File);
            pdf2 = PDDocument.load(pdf2File);
            outPdf = new PDDocument();

            //Create output PDF Frame

            PDRectangle pdf1Frame = pdf1.getPage(0).getCropBox();
            PDRectangle pdf2Frame = pdf2.getPage(0).getCropBox();

            PDRectangle outPdfFrame = new PDRectangle(pdf1Frame.getWidth() + pdf2Frame.getWidth(), Math.max(pdf1Frame.getHeight(), pdf2Frame.getHeight()));

            // Create output page  with calculated frame and add it to the document
            COSDictionary dict = new COSDictionary();
            dict.setItem(COSName.TYPE, COSName.PAGE);
            dict.setItem(COSName.MEDIA_BOX, outPdfFrame);
            dict.setItem(COSName.CROP_BOX, outPdfFrame);
            dict.setItem(COSName.ART_BOX, outPdfFrame);
            PDPage outPdfPage = new PDPage(dict);
            outPdf.addPage(outPdfPage);

            // Source PDF pages has to be imported as form XObjects to be able to insert them at a specific point in the output page
            LayerUtility layerUtility = new LayerUtility(outPdf);
            PDFormXObject formPdf1 = layerUtility.importPageAsForm(pdf1, 0);
            PDFormXObject formPdf2 = layerUtility.importPageAsForm(pdf2, 0);

            // Add form objects to output page
            AffineTransform afLeft = new AffineTransform();
            layerUtility.appendFormAsLayer(outPdfPage, formPdf1, afLeft, layerName1);
            AffineTransform afRight = AffineTransform.getTranslateInstance(pdf1Frame.getWidth(), 0.0);
            layerUtility.appendFormAsLayer(outPdfPage, formPdf2, afRight, layerName2);
            outPdf.save(outPdfFile + "//" + mergedFileName);

            pdf1File.delete();
            pdf2File.delete();

        } finally {
            try {
                if (pdf1 != null) pdf1.close();
                if (pdf2 != null) pdf2.close();
                if (outPdf != null) outPdf.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    private void ignorePagesByTitle(String prodReport, String stageReport, String[] titleArray,
                                    List<PageModel> prodPages, List<PageModel> stagePages) throws IOException {
        int totalPagesInProd = PDFUtilHelper.getPageCount(prodReport);
        int totalPagesInStage = PDFUtilHelper.getPageCount(stageReport);

        for (int i = 1; i <= totalPagesInProd; i++) {
            String prodPageTitle = getPageTitle(prodReport, i);
//            System.out.println( i + " page tile is:" + prodPageTitle);
            if (isPageIgnored(prodPageTitle, titleArray)) {
//                System.out.println("Page " + i + " in Production contains " + prodPageTitle);
            } else {
                String prodPageText = pdfCompareHelper.getPageTextFromPDF(prodReport, i, i);
                PageModel page = new PageModel();
                page.setPageText(prodPageText);
                page.setPageTitle(prodPageTitle);
                page.setPageNumber(i);
                prodPages.add(page);
            }
        }

        for (int i = 1; i <= totalPagesInStage; i++) {
            String stagePageTitle = getPageTitle(stageReport, i);
            if (isPageIgnored(stagePageTitle, titleArray)) {
//                System.out.println("Page " + i + " in Stage contains " + stagePageTitle);
            } else {
                String stageText = pdfCompareHelper.getPageTextFromPDF(stageReport, i, i);
                PageModel page = new PageModel();
                page.setPageText(stageText);
                page.setPageTitle(stagePageTitle);
                page.setPageNumber(i);
                stagePages.add(page);
            }
        }
    }

    private void setPagesValues(String prodReport, String stageReport,
                                List<PageModel> prodPages, List<PageModel> stagePages) throws IOException {
        int totalPagesInProd = PDFUtilHelper.getPageCount(prodReport);
        int totalPagesInStage = PDFUtilHelper.getPageCount(stageReport);

        for (int i = 0; i <= totalPagesInProd; i++) {
            String prodPageText = pdfCompareHelper.getPageTextFromPDF(prodReport, i, i);
            PageModel page = new PageModel();
            page.setPageText(prodPageText);
            page.setPageNumber(i);
            prodPages.add(page);
        }

        for (int i = 1; i <= totalPagesInStage; i++) {
            String stageText = pdfCompareHelper.getPageTextFromPDF(stageReport, i, i);
            PageModel page = new PageModel();
            page.setPageText(stageText);
            page.setPageNumber(i);
            stagePages.add(page);
        }
    }

    private String getIgnoredTitles() throws IOException {
        String ignoredTitlesFile = "files//IgnoredTitles.txt";
        String titles = "";
        BufferedReader br = new BufferedReader(new FileReader(ignoredTitlesFile));
        try {
            StringBuilder sb = new StringBuilder();
            String line = br.readLine();
            while (line != null) {
                sb.append(line);
                sb.append(System.lineSeparator());
                line = br.readLine();
            }
            titles = sb.toString();
        } finally {
            br.close();
        }
        return titles;
    }

    public static void deleteTempFiles() throws IOException {
        String finalReportPath = "files//Report//";
        String finalResultPath = "files//Results//";
        String temp = "files//prod//temp//";
        String directory = System.getProperty("user.dir");
        File folder = new File(directory);
        Arrays.stream(Objects.requireNonNull(folder.listFiles())).filter(f -> f.getName().endsWith(".png"))
                .forEach(File::delete);
        Arrays.stream(Objects.requireNonNull(folder.listFiles())).filter(f -> f.getName().endsWith(".pdf"))
                .forEach(File::delete);
        FileUtils.cleanDirectory(new File(finalResultPath));
        FileUtils.cleanDirectory(new File(finalReportPath));
        FileUtils.cleanDirectory(new File(temp));

    }

    public static LinkedList<String> getFileNameFromFolder(String path) {
        File stageFolder = new File(path);
        File[] listOfStageFiles = stageFolder.listFiles();
        System.out.println("Files present inside folder : " + path);
        LinkedList<String> fileName = new LinkedList<>();
        for (int stageCount = 0; stageCount < listOfStageFiles.length; stageCount++) {
            if (listOfStageFiles[stageCount].isFile()) {
                System.out.println(stageCount + " : " + listOfStageFiles[stageCount].getName());
                fileName.add(listOfStageFiles[stageCount].getName());
            }
        }
        return fileName;
    }

    public void getFileNameFromFolder() {
        String stageFilePath = "files//Stage//";
        String prodFilePath = "files//Prod//";

        File stageFolder = new File(stageFilePath);
        File[] listOfStageFiles = stageFolder.listFiles();
        System.out.println("Files present inside Stage folder : ");
        for (int stageCount = 0; stageCount < listOfStageFiles.length; stageCount++) {
            if (listOfStageFiles[stageCount].isFile()) {
                System.out.println(stageCount + " : " + listOfStageFiles[stageCount].getName());
                stageFileNames.add(listOfStageFiles[stageCount].getName());
            }
        }
        File prodFolder = new File(prodFilePath);
        File[] listOfProdFiles = prodFolder.listFiles();
        System.out.println("Files present inside Prod folder : ");
        for (int prodCount = 0; prodCount < listOfProdFiles.length; prodCount++) {
            if (listOfProdFiles[prodCount].isFile()) {
                System.out.println(prodCount + " : " + listOfProdFiles[prodCount].getName());
                prodFileNames.add(listOfProdFiles[prodCount].getName());
            }
        }
    }

    private String getPageTitle(String report, int i) throws IOException {
        String pageTitle = "";
        pageTitle = pdfCompareHelper.getSpecificTextFromPDF(report, i, i);
        return pageTitle.trim();
    }

    private boolean isPageIgnored(String pageTitle, String[] titleArray) {
        pageTitle = pageTitle.trim();
        for (String title : titleArray) {
            title = title.trim();
            if (StringUtils.isNotEmpty(title) && pageTitle.contains(title)) {
                return true;
            }
        }
        return false;
    }

}
