package util;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.ArrayList;

import javax.imageio.ImageIO;

import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.graphics.PDXObject;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.pdfbox.tools.imageio.ImageIOUtil;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.commons.io.FileUtils;


public class PDFUtilHelper {
    private String imageDestinationPath;
    private boolean bTrimWhiteSpace;
    private boolean bHighlightPdfDifference;
    private Color imgColor;
    private PDFTextStripper stripper;
    private boolean bCompareAllPages;
    private CompareMode compareMode;
    private String[] excludePattern;
    private int startPage = 1;
    private int endPage = -1;

    public PDFUtilHelper() {
        this.imgColor = Color.MAGENTA;
        this.bCompareAllPages = false;
        this.compareMode = CompareMode.TEXT_MODE;
        System.setProperty("sun.java2d.cmm", "sun.java2d.cmm.kcms.KcmsServiceProvider");
    }


    public void setCompareMode(CompareMode mode) {
        this.compareMode = mode;
    }


    public String getImageDestinationPath() {
        return this.imageDestinationPath;
    }

    public static int getPageCount(String file) throws IOException {
        PDDocument doc = PDDocument.load(new File(file));
        int pageCount = doc.getNumberOfPages();
        doc.close();
        return pageCount;
    }

    public String getText(String file, int startPage, int endPage) throws IOException {
        return this.getPDFText(file, startPage, endPage);
    }

    private String getPDFText(String file, int startPage, int endPage) throws IOException {

        PDDocument doc = PDDocument.load(new File(file));

        PDFTextStripper localStripper = new PDFTextStripper();
        if (null != this.stripper) {
            localStripper = this.stripper;
        }

        this.updateStartAndEndPages(file, startPage, endPage);
        localStripper.setStartPage(this.startPage);
        localStripper.setEndPage(this.endPage);

        String txt = localStripper.getText(doc);
        if (this.bTrimWhiteSpace) {
            txt = txt.trim().replaceAll("\\s+", " ").trim();
        }

        doc.close();
        return txt;
    }


    public void excludeText(String... regexs) {
        this.excludePattern = regexs;
    }

    public boolean compare(String file1, String file2) throws IOException {
        return this.comparePdfFiles(file1, file2, -1, -1);
    }

    private boolean comparePdfFiles(String file1, String file2, int startPage, int endPage) throws IOException {
        if (CompareMode.TEXT_MODE == this.compareMode)
            return comparepdfFilesWithTextMode(file1, file2, startPage, endPage);
        else
            return comparePdfByImage(file1, file2, startPage, endPage);
    }

    private boolean comparepdfFilesWithTextMode(String file1, String file2, int startPage, int endPage) throws IOException {

        String file1Txt = this.getPDFText(file1, startPage, endPage).trim();
        String file2Txt = this.getPDFText(file2, startPage, endPage).trim();

        if (null != this.excludePattern && this.excludePattern.length > 0) {
            for (int i = 0; i < this.excludePattern.length; i++) {
                file1Txt = file1Txt.replaceAll(this.excludePattern[i], "");
                file2Txt = file2Txt.replaceAll(this.excludePattern[i], "");
            }
        }
        System.out.println("File 1 text :" + file1Txt);
        System.out.println("File 2 text :" + file2Txt);
        boolean result = file1Txt.equalsIgnoreCase(file2Txt);

        if (!result) {
            System.out.println("PDF content does not match");
        }

        return result;
    }


    public List<String> savePdfAsImage(String file, int startPage, int endPage, String fileName) throws IOException {
        return this.saveAsImage(file, startPage, endPage, fileName);
    }


    private List<String> saveAsImage(String file, int startPage, int endPage, String fileNameExtn) throws IOException {
        ArrayList<String> imgNames = new ArrayList<String>();

        try {
            File sourceFile = new File(file);
            this.createImageDestinationDirectory(file);
            this.updateStartAndEndPages(file, startPage, endPage);

            String fileName = sourceFile.getName().replace(".pdf", "");

            PDDocument document = PDDocument.load(sourceFile);
            PDFRenderer pdfRenderer = new PDFRenderer(document);
            for (int iPage = this.startPage - 1; iPage < this.endPage; iPage++) {
                String fname = this.imageDestinationPath + fileName + "_" + (iPage + 1) + fileNameExtn + ".png";
                BufferedImage image = pdfRenderer.renderImageWithDPI(iPage, 300, ImageType.RGB);
                ImageIOUtil.writeImage(image, fname, 300);
                imgNames.add(fname);
            }
            document.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return imgNames;
    }

    public boolean compare(String file1, String file2, int startPage, int endPage, boolean highlightImageDifferences, boolean showAllDifferences) throws IOException {
        this.compareMode = CompareMode.VISUAL_MODE;
        this.bHighlightPdfDifference = highlightImageDifferences;
        this.bCompareAllPages = showAllDifferences;
        return this.comparePdfByImage(file1, file2, startPage, endPage);
    }


    private boolean comparePdfByImage(String file1, String file2, int startPage, int endPage) throws IOException {
        System.out.println("file1 : " + file1);
        System.out.println("file2 : " + file2);
        this.createImageDestinationDirectory(file2);
        this.updateStartAndEndPages(file1, startPage, endPage);
        return this.convertToImageAndCompare(file1, file2, this.startPage, this.endPage);
    }

    private boolean convertToImageAndCompare(String file1, String file2, int startPage, int endPage) throws IOException {
        boolean result = true;
        PDDocument doc1 = null;
        PDDocument doc2 = null;
        PDFRenderer pdfRenderer1 = null;
        PDFRenderer pdfRenderer2 = null;

        try {

            doc1 = PDDocument.load(new File(file1));
            doc2 = PDDocument.load(new File(file2));

            pdfRenderer1 = new PDFRenderer(doc1);
            pdfRenderer2 = new PDFRenderer(doc2);


            for (int iPage = startPage - 1; iPage < endPage; iPage++) {
                String fileName = new File(file1).getName().replace(".pdf", "_") + (iPage + 1);
                fileName = this.getImageDestinationPath() + "/" + fileName + "_diff.png";

                System.out.println("Comparing Page No : " + (iPage + 1));
                BufferedImage image1 = pdfRenderer1.renderImageWithDPI(iPage, 300, ImageType.RGB);
                BufferedImage image2 = pdfRenderer2.renderImageWithDPI(iPage, 300, ImageType.RGB);
                result = ImageUtil.compareAndHighlight(image1, image2, fileName, this.bHighlightPdfDifference, this.imgColor.getRGB()) && result;
                if (!this.bCompareAllPages && !result) {
                    break;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            doc1.close();
            doc2.close();
        }
        return result;
    }


    public List<String> extractImages(String file, int startPage, int endPage) throws IOException {
        return this.extractimages(file, startPage, endPage);
    }


    /**
     * This method extracts all the embedded images of the pdf document
     */
    private List<String> extractimages(String file, int startPage, int endPage) {
        ArrayList<String> imgNames = new ArrayList<String>();
        boolean bImageFound = false;
        try {

            this.createImageDestinationDirectory(file);
            String fileName = this.getFileName(file).replace(".pdf", "_resource");

            PDDocument document = PDDocument.load(new File(file));
            PDPageTree list = document.getPages();

            this.updateStartAndEndPages(file, startPage, endPage);

            int totalImages = 1;
            for (int iPage = this.startPage - 1; iPage < this.endPage; iPage++) {
                System.out.println("Page No : " + (iPage + 1));
                PDResources pdResources = list.get(iPage).getResources();
                for (COSName c : pdResources.getXObjectNames()) {
                    PDXObject o = pdResources.getXObject(c);
                    if (o instanceof org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject) {
                        bImageFound = true;
                        String fname = this.imageDestinationPath + "/" + fileName + "_" + totalImages + ".png";
                        ImageIO.write(((org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject) o).getImage(), "png", new File(fname));
                        imgNames.add(fname);
                        totalImages++;
                    }
                }
            }
            document.close();
            if (bImageFound)
                System.out.println("Images are saved @ " + this.imageDestinationPath);
            else
                System.out.println("No images were found in the PDF");
        } catch (Exception e) {
            e.printStackTrace();
        }
        return imgNames;
    }

    private void createImageDestinationDirectory(String file) throws IOException {
        if (null == this.imageDestinationPath) {
            File sourceFile = new File(file);
            String destinationDir = sourceFile.getParent() + "/temp/";
            this.imageDestinationPath = destinationDir;
            this.createFolder(destinationDir);
        }
    }

    private boolean createFolder(String dir) throws IOException {
        FileUtils.deleteDirectory(new File(dir));
        return new File(dir).mkdir();
    }

    private String getFileName(String file) {
        return new File(file).getName();
    }

    private void updateStartAndEndPages(String file, int start, int end) throws IOException {

        PDDocument document = PDDocument.load(new File(file));
        int pagecount = document.getNumberOfPages();
        if ((start > 0 && start <= pagecount)) {
            this.startPage = start;
        } else {
            this.startPage = 1;
        }
        if ((end > 0 && end >= start && end <= pagecount)) {
            this.endPage = end;
        } else {
            this.endPage = pagecount;
        }
        document.close();
    }

    public String getPdfPageText(String file, int startPage, int endPage) throws IOException {
        PDDocument doc = PDDocument.load(new File(file));
        PDFTextStripperByArea pdfTextStripperByArea = new PDFTextStripperByArea();
        this.updateStartAndEndPages(file,startPage,endPage);
        pdfTextStripperByArea.setStartPage(this.startPage);
        pdfTextStripperByArea.setEndPage(this.endPage);
        Rectangle2D rectangle2D = new Rectangle2D.Float(60,20,1000,50);
        pdfTextStripperByArea.addRegion("region", rectangle2D);
        PDPage pdPage = doc.getPage(startPage-1);
        pdfTextStripperByArea.extractRegions(pdPage);
        String textFromRegion = pdfTextStripperByArea.getTextForRegion("region");
        doc.close();;
        return textFromRegion;
    }
}
