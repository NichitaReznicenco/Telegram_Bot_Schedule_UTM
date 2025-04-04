package org.yourcompany.yourproject;
import com.spire.pdf.FileFormat;  
import com.spire.pdf.PdfDocument;  

public class JavaSmt {

    public static void main(String[] args) {
        System.out.println("Hello World!");

         //Create a PdfDocument instance  
        PdfDocument pdf = new PdfDocument();  
        //Load a PDF file  
        pdf.loadFromFile("pathToPdf");  
        //Save to .docx file  
        pdf.saveToFile("fileName.xlsx", FileFormat.XLSX);  
        pdf.close();  
    }
}
