package com.example.demo.controller;

import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;

@Controller
public class PDFController {

    @GetMapping("/")
    public String home(){
        return "fileupload";
    }

    @PostMapping("/upload")
    public String pdfToExcel(@RequestParam("file") MultipartFile file) throws IOException {
        System.out.println(file.getOriginalFilename());
        PDDocument pdf = PDDocument.load((File) file);
        PDFTextStripper pdfTextStripper = new PDFTextStripper();
        PDPageTree pages = pdf.getPages();
        pdfTextStripper.setStartPage(0);
        pdfTextStripper.setEndPage(pdf.getNumberOfPages());
        String pdfText = pdfTextStripper.getText(pdf);
        String[] lines = pdfText.split(System.lineSeparator());
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("sheet");
        int rowNo = 0;
        for (String line : lines) {
            Row row = sheet.createRow(rowNo++);
            String[] cells = line.split("\\t");
            int cellNum = 0;
            for (String cellValue : cells) {
                Cell cell = row.createCell(cellNum++);
                cell.setCellValue(cellValue);
            }

            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            workbook.close();

            byte[] excelBytes = outputStream.toByteArray();
            System.out.println(line);
        }

        pdf.close();

        return "index";
    }
}
