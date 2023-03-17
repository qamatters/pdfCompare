package helper;

import net.sourceforge.tess4j.Tesseract;
import org.apache.commons.io.comparator.LastModifiedFileComparator;
import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import util.PDFUtilHelper;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;

public class imageCompareHelper {

    public static void createPngImage(BufferedImage image, String fileName) throws IOException {
        ImageIO.write(image, "png", new File(fileName));
    }

    public static void createJpgImage(BufferedImage image, String fileName) throws IOException {
        ImageIO.write(image, "jpg", new File(fileName));
    }

    public static String compareWithBaseImage(File baseImage, File CompareImage, String resultOfComparison) throws IOException {
        BufferedImage bImage = ImageIO.read(baseImage);
        BufferedImage cImage = ImageIO.read(CompareImage);
        int height = bImage.getHeight();
        int width = bImage.getWidth();
        int highlight = Color.MAGENTA.getRGB();

        if (baseImage.toString().contains("CRP") == true ||
                baseImage.toString().contains("crp")
                || baseImage.toString().contains("PRP")
                || baseImage.toString().contains("prp")
                || baseImage.toString().contains("IRP")
                || baseImage.toString().contains("irp")) {
            width = width - 230; // Ignore page numbers from PRP/CRP reports
        }
        String diffInd = "N";
        BufferedImage rImage = new BufferedImage(width, height, BufferedImage.TYPE_4BYTE_ABGR);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                try {
                    int pixelC = cImage.getRGB(x, y);
                    int pixelB = bImage.getRGB(x, y);
                    if (pixelB == pixelC) {
                        rImage.setRGB(x, y, bImage.getRGB(x, y));
                    } else {
                        diffInd = "Y";
//                        int a = 0xff | bImage.getRGB(x, y) >> 24,
//                                r = 0xff & bImage.getRGB(x, y) >> 16,
//                                g = 0xff & bImage.getRGB(x, y) >> 8,
//                                b = 0x00 & bImage.getRGB(x, y);
//                        int modifiedRGB = a << 24 | r << 16 | g << 8 | b;
                            rImage.setRGB(x, y,highlight);
                    }
                } catch (Exception e) {
                    rImage.setRGB(x, y, 0x80ff0000);
                }
            }
        }
        String filePath = baseImage.toPath().toString();
        String fileExtension = filePath.substring(filePath.lastIndexOf('.'), filePath.length());
        if (fileExtension.toUpperCase().contains("PNG")) {
            createPngImage(rImage, resultOfComparison + fileExtension);
        } else {
            createJpgImage(rImage, resultOfComparison + fileExtension);
        }

        if (diffInd.equals("Y")) {
            return resultOfComparison + fileExtension;
        } else {
            return diffInd;
        }
    }

    public static void createPDFFromSingleFile(String diffOutPUtFileNames, String pdfName) throws IOException {
        PDDocument pdfDoc = new PDDocument();
        try{
            InputStream inputStream = new FileInputStream(diffOutPUtFileNames);
            BufferedImage bImage = ImageIO.read(inputStream);
            float width = bImage.getWidth();
            float height = bImage.getHeight();

            PDPage page = new PDPage(new PDRectangle(width, height));

            pdfDoc.addPage(page);
            PDImageXObject pdImage = PDImageXObject.createFromFile(diffOutPUtFileNames,pdfDoc);
            PDPageContentStream contentStream = new PDPageContentStream(
                    pdfDoc,page, PDPageContentStream.AppendMode.APPEND,true,true
            );
            contentStream.drawImage(pdImage, 20f, 20f);
            contentStream.close();
            inputStream.close();
            pdfDoc.save(pdfName);
        } finally {
            pdfDoc.close();
        }
    }
}
