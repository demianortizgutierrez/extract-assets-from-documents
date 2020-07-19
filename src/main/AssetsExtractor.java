package main;

import main.lib.*;

public class AssetsExtractor {

    public static void main(String[] args) throws Exception {

        Extractor PPTXExtractor = new PPTXExtractor();
        String PPTXFileName = "sample.pptx";
        PPTXExtractor.extract(PPTXFileName);
        
        Extractor DocExtractor = new DocExtractor();
        String DocFileName = "sample.docx";
        DocExtractor.extract(DocFileName);
        
        Extractor PDFExtractor = new PDFExtractor();
        String PDFFileName = "sample.pdf";
        PDFExtractor.extract(PDFFileName);

    }

}
