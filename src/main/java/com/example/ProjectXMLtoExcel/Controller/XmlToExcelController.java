package com.example.ProjectXMLtoExcel.Controller;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;


import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;

@RestController
public class XmlToExcelController {

    @GetMapping("/convert")
    public void convertXmlToExcel(HttpServletResponse response) throws Exception {
        // Read input XML file into a Workbook
        String inputFile = "src/main/resources/data.xml";
        Workbook workbook = new XSSFWorkbook();

        // Get the first sheet in the workbook
        Sheet sheet = workbook.createSheet("Employees");

        // Create headers row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Name");
        headerRow.createCell(1).setCellValue("Email");
        headerRow.createCell(2).setCellValue("Age");
        headerRow.createCell(3).setCellValue("Salary");

        // Create a DOM parser factory
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();

        // Create a DOM parser
        DocumentBuilder builder = factory.newDocumentBuilder();

        // Parse the input XML file into a DOM document

        Document document = builder.parse(new File(inputFile));

        // Get the "employee" tags from the document
        NodeList employeeList = document.getElementsByTagName("employee");

        // Populate data rows from XML
        int rowIndex = 1;
        for (int i = 0; i < employeeList.getLength(); i++) {
            Element employee = (Element) employeeList.item(i);
            String name = employee.getAttribute("name");
            String email = employee.getAttribute("email");
            int age = Integer.parseInt(employee.getAttribute("age"));
            int salary = Integer.parseInt(employee.getAttribute("salary"));

            // Populate Excel data row with employee attributes
            Row dataRow = sheet.createRow(rowIndex++);
            dataRow.createCell(0).setCellValue(name);
            dataRow.createCell(1).setCellValue(email);
            dataRow.createCell(2).setCellValue(age);
            dataRow.createCell(3).setCellValue(salary);
        }

        // Write output Excel file to local file system
        String outputFile = "src/main/resources/converted.xlsx";
        FileOutputStream outputStream = new FileOutputStream(outputFile);
        workbook.write(outputStream);
        outputStream.close();

        // Set headers for Excel file download
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=converted.xlsx");
        response.setHeader("Cache-Control", "no-cache");
        response.setHeader("Pragma", "no-cache");

        // Write output
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        workbook.write(baos);
        response.getOutputStream().write(baos.toByteArray());
        response.getOutputStream().flush();
        workbook.close();
    }


}
