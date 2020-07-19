package main.lib;

import java.awt.Dimension;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class PPTXExtractor implements Extractor {

    public void extract(String fileName) {

        String file = "";
        PrintStream out = System.out;
        ClassLoader classLoader = getClass().getClassLoader();

        if (classLoader.getResource(fileName) != null) {
            file = classLoader.getResource(fileName).getFile();
        } else {
            file = "resources/" + fileName;
        }

        try {
            FileInputStream is = new FileInputStream(file);
            XMLSlideShow ppt = new XMLSlideShow(is);
            is.close();
    
            // Get the document's embedded files.
            for (PackagePart p : ppt.getAllEmbeddedParts()) {
                String type = p.getContentType();
                // typically file name
                String name = p.getPartName().getName();
                out.println("Embedded file (" + type + "): " + name);
    
                InputStream pIs = p.getInputStream();
                // make sense of the part data
                pIs.close();
    
            }
    
            // Get the document's embedded files.
            for (XSLFPictureData data : ppt.getPictureData()) {
                String type = data.getContentType();
                String name = data.getFileName();
                out.println("Picture (" + type + "): " + name);
    
                InputStream pIs = data.getInputStream();
                // make sense of the image data
                pIs.close();
            }
    
            // size of the canvas in points
            Dimension pageSize = ppt.getPageSize();
            out.println("Pagesize: " + pageSize);
    
            for (XSLFSlide slide : ppt.getSlides()) {
                for (XSLFShape shape : slide) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape txShape = (XSLFTextShape) shape;
                        out.println(txShape.getText());
                    } else if (shape instanceof XSLFPictureShape) {
                        XSLFPictureShape pShape = (XSLFPictureShape) shape;
                        XSLFPictureData pData = pShape.getPictureData();
                        out.println(pData.getFileName());
                    } else {
                        out.println("Process me: " + shape.getClass());
                    }
                }
            }
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    @Override
    public void showFiles() {
        // TODO Auto-generated method stub
        
    }
}
