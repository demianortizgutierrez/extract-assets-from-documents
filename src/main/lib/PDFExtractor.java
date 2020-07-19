package main.lib;

import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.graphics.PDXObject;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;

import javax.imageio.ImageIO;
import java.io.File;
import java.io.IOException;

public class PDFExtractor implements Extractor {

    private static final String OUTPUT_DIR = "resources/";

    public void extract(String fileName) {
        
        String file = "";

        ClassLoader classLoader = getClass().getClassLoader();

        if (classLoader.getResource(fileName) != null) {
            file = classLoader.getResource(fileName).getFile();
        } else {
            file = OUTPUT_DIR + fileName;
        }

        try {
            final PDDocument document = PDDocument.load(new File(file));
            PDPageTree list = document.getPages();
            for (PDPage page : list) {
                PDResources pdResources = page.getResources();
                int i = 1;
                for (COSName name : pdResources.getXObjectNames()) {
                    PDXObject o = pdResources.getXObject(name);
                    if (o instanceof PDImageXObject) {
                        PDImageXObject image = (PDImageXObject) o;
                        String filename = OUTPUT_DIR + "extracted-image-" + i + ".png";
                        ImageIO.write(image.getImage(), "png", new File(filename));
                        i++;
                    }
                }
            }

        } catch (IOException e) {
            System.err.println("Exception while trying to load PDF document - " + e);
        }
    }

    public void showFiles() {

    }
}
