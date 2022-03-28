package it.devlec.excel;

import it.devlec.csv.EsempioCSV;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVRecord;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URISyntaxException;
import java.nio.file.Paths;
import java.util.Iterator;

public class EsempioExcel {
    private static final Logger logger = LogManager.getLogger(EsempioExcel.class);

    public Workbook leggiExcel(String filename) {
        logger.debug("Provo a leggere un file excel");
        String excelDiProva = null;
        try {
            excelDiProva = Paths.get(ClassLoader.getSystemResource(filename)
                    .toURI()).toString();
        } catch (URISyntaxException e) {
            logger.error("Errore nel trovare nel creare il file");
        }
        FileInputStream file = null;
        try {
            file = new FileInputStream(excelDiProva);
            Workbook workbook = new XSSFWorkbook(file);
            workbook.close();
            IOUtils.closeQuietly(file);
            return workbook;
        } catch (IOException e) {
            logger.error("Errore nel leggere il mio excel", e);
        }
        return null;
    }


    public void testLetturaExcel() {
        logger.debug("Provo a leggere un file excel");
        Workbook workbook = leggiExcel("excel.xlsx");
        Sheet sheet = workbook.getSheetAt(0);
        try {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    logger.info("Valore " + cell.getStringCellValue());
                }
            }
            workbook.close();

        } catch (IOException e) {
            logger.error("Errore nel leggere il mio excel", e);
        }

    }

    public void scriviIlMioFileExcel() {
        String excelFile = EsempioCSV.getFilePath("excel.xlsx");
        File parent = new File(excelFile).getParentFile();
        String mioExcel = parent.getAbsolutePath() + File.separator + "mioExcelGenerato.xlsx";
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Persona");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);

        Row header = sheet.createRow(0);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Nome");
        headerCell.setCellStyle(headerStyle);
        headerCell = header.createCell(1);
        headerCell.setCellValue("Eta");
        headerCell.setCellStyle(headerStyle);

        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);

        Row row = sheet.createRow(2);
        Cell cell = row.createCell(0);
        cell.setCellValue("Mario Rossi");
        cell.setCellStyle(style);

        cell = row.createCell(1);
        cell.setCellValue(20);
        cell.setCellStyle(style);


        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(mioExcel);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        IOUtils.closeQuietly(outputStream);
    }
    public void leggiDaCSV() {
        String csvFile = EsempioCSV.getFilePath("esempio.csv");
        File parent = new File(csvFile).getParentFile();
        String mioExcel = parent.getAbsolutePath() + File.separator + "autori.xlsx";
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Autori");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setWrapText(true);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setFont(font);

        int rowIndex = 0;

        try {
            Reader  in = new FileReader(csvFile);
            Iterable<CSVRecord> records = CSVFormat.DEFAULT.builder().build().parse(in);
            for (Row row : sheet) {
                for(Cell cella : row) {
                    if(rowIndex == 0) {
                        cella.setCellStyle(headerStyle);
                    }
                    Iterator<CSVRecord> csv = records.iterator();
                    if(csv.hasNext()) {
                        String valoreCella = csv.next().toString();
                        cella.setCellValue(valoreCella);
                    }
                }
                rowIndex++;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(mioExcel);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        IOUtils.closeQuietly(outputStream);
    }
}
