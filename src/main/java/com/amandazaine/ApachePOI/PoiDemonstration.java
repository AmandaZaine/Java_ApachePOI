package com.amandazaine.ApachePOI;

import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

public class PoiDemonstration {

    public static void saveFile() throws IOException {
        //Cria uma planilha de Excel
        Workbook workbook = new XSSFWorkbook();

        //Cria uma aba na planilha
        Sheet sheet1 = workbook.createSheet("Sheet 1");

        Row row = sheet1.createRow(0); // Cria uma linha
        Cell cell = row.createCell(0); // Cria uma célula na primeira coluna
        cell.setCellValue("Hello, Excel!"); // Define o valor da célula


        //Local onde o arquivo será salvo
        String filePath = "/home/amanda/Documentos/myfile.xlsx";

        //Cria um objeto File que representa o arquivo no endereço indicado
        File myFile = new File(filePath);

        //Cria um -fluxo de saída- para gravar os dados no arquivo representado pelo objeto myFile
        FileOutputStream fileOutputStream = new FileOutputStream(myFile);

        //Escreve os dados no fluxo de saída associado ao arquivo
        //Fluxo de saída é o canal usado para enviar dados para um destino externo.
        workbook.write(fileOutputStream);

        //Fecha o FileOutputStream, liberando os recursos alocados.
        fileOutputStream.close();
        System.out.println("File created");
    }

    public static void readFile() throws IOException {
        String filePath = "/home/amanda/Documentos/myfile.xlsx";
        File myFile = new File(filePath);

        FileInputStream fileInputStream = null;

        if(myFile.isFile() && myFile.exists()) {
            fileInputStream = new FileInputStream(myFile);
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(0);
            Cell cell = row.getCell(0);

            System.out.println("File open successfully: " + cell.getStringCellValue());
        } else {
            System.out.println("File not open");
        }

        if(fileInputStream != null) {
            fileInputStream.close();
        }
    }

    public static void fillInSheet(List<List<String>> content) throws IOException {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("Sheet 1");

        for (int i = 0; i < content.size(); i++) {
            Row row = sheet1.createRow(i);

            for (int j = 0; j < content.get(i).size() ; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(content.get(i).get(j));
            }
        }

        String filePath = "/home/amanda/Documentos/myfile.xlsx";
        File myFile = new File(filePath);

        FileOutputStream fileOutputStream = new FileOutputStream(myFile);
        workbook.write(fileOutputStream);

        fileOutputStream.close();
        System.out.println("File created");
    }

    public static void formatDateInACell(List<List<String>> content) throws IOException, ParseException {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("Sheet 1");

        CellStyle cellStyle = workbook.createCellStyle();

        DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat("dd/MM/yyyy"));

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cellStyle.setBorderBottom(BorderStyle.DOTTED);
        cellStyle.setBottomBorderColor(IndexedColors.PINK.getIndex());
        cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        cellStyle.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.FINE_DOTS);


        for (int i = 0; i < content.size(); i++) {
            Row row = sheet1.createRow(i);

            for (int j = 0; j < content.get(i).size() ; j++) {
                Cell cell = row.createCell(j);
                SimpleDateFormat formatoData = new SimpleDateFormat("dd/MM/yyyy");
                cell.setCellValue(formatoData.parse(content.get(i).get(j)));
                cell.setCellStyle(cellStyle);
            }
        }

        String filePath = "/home/amanda/Documentos/myfile.xlsx";
        File myFile = new File(filePath);

        FileOutputStream fileOutputStream = new FileOutputStream(myFile);
        workbook.write(fileOutputStream);

        fileOutputStream.close();
        System.out.println("File created");
    }

    public static void cellMerge(List<List<String>> content) throws IOException {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("Sheet 1");

        for (int i = 0; i < content.size(); i++) {
            Row row = sheet1.createRow(i);

            for (int j = 0; j < content.get(i).size() ; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(content.get(i).get(j));
            }
        }

        Row mergedRow = sheet1.getRow(1);
        Cell mergedCell = mergedRow.createCell(content.getFirst().size() + 2);
        mergedCell.setCellValue("Células mescladas!");

        CellStyle cellStyle = workbook.createCellStyle();

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle.setBorderTop(BorderStyle.MEDIUM);
        cellStyle.setBorderLeft(BorderStyle.MEDIUM);
        cellStyle.setBorderRight(BorderStyle.MEDIUM);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());

        cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight(Short.parseShort("500"));
        //font.setFontHeightInPoints((short) 15);
        cellStyle.setFont(font);

        mergedCell.setCellStyle(cellStyle);

        //Mesclar células
        sheet1.addMergedRegion(
                new CellRangeAddress(
                        1,
                        10,
                        content.getFirst().size() + 2,
                        content.getFirst().size() + 5)
        );

        String filePath = "/home/amanda/Documentos/myfile.xlsx";
        File myFile = new File(filePath);

        FileOutputStream fileOutputStream = new FileOutputStream(myFile);
        workbook.write(fileOutputStream);

        fileOutputStream.close();
        System.out.println("File created");
    }

    public static void exercicio1() throws IOException, DecoderException {

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("MySheet");
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);

        //Configurando estilo da célula
        XSSFCellStyle cellStyle = workbook.createCellStyle();

        //Definindo cor de fundo
        XSSFColor verdeClaro = criarCor("42f551");
        cellStyle.setFillForegroundColor(verdeClaro);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //Definindo fonte
        XSSFFont fonte = workbook.createFont();
        fonte.setFontName("Montserrat Black");
        fonte.setItalic(true);
        fonte.setFontHeightInPoints((short) 25);
        fonte.setColor(IndexedColors.DARK_GREEN.getIndex());
        cellStyle.setFont(fonte);

        //Definindo alinhamento do texto na celula
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cell.setCellValue("Aprendendo Apache POI");
        cell.setCellStyle(cellStyle);

        //Configurando estilo da coluna
        sheet.autoSizeColumn(0);

        OutputStream outputStream = null;

        try {
            //Salva o conteúdo de um objeto do tipo Workbook em um arquivo Excel
            outputStream = new FileOutputStream("/home/amanda/Documentos/novoarquivo.xlsx");
            workbook.write(outputStream);
            System.out.println("Arquivo Excel criado!");
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            workbook.close();
            outputStream.close();
        }

    }

    private static XSSFColor criarCor(String corHexadecimal) throws DecoderException {
        try {
            byte[] rgb = Hex.decodeHex(corHexadecimal);
            return new XSSFColor(rgb);
        } catch (DecoderException e) {
            e.printStackTrace();
            throw new RuntimeException("Erro ao criar cor");
        }
    }

    public static void exercicio2() throws IOException, DecoderException {

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("MySheet");
        XSSFRow row = sheet.createRow(2);
        XSSFCell cell = row.createCell(2);

        //Configurando estilo da célula
        XSSFCellStyle cellStyle = workbook.createCellStyle();

        //Definindo alinhamento do texto na celula
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        //Mesclando celulas
        CellRangeAddress mesclagem = new CellRangeAddress(2,10,2,10);

        cell.setCellValue("Aprendendo Apache POI");
        cell.setCellStyle(cellStyle);

        //Configurando estilo da coluna
        sheet.autoSizeColumn(0);
        sheet.addMergedRegion(mesclagem);

        OutputStream outputStream = null;

        try {
            //Salva o conteúdo de um objeto do tipo Workbook em um arquivo Excel
            outputStream = new FileOutputStream("/home/amanda/Documentos/novoarquivo.xlsx");
            workbook.write(outputStream);
            System.out.println("Arquivo Excel criado!");
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            workbook.close();
            outputStream.close();
        }

    }

    public static void exercicio3() throws IOException, DecoderException {

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("MySheet");
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell0 = row.createCell(0);
        XSSFCell cell1 = row.createCell(1);

        //Configurando o formato do dado da celula
        XSSFCellStyle cellStyle0 = workbook.createCellStyle();
        cellStyle0.setDataFormat(workbook.createDataFormat().getFormat("dd/MM/yyyy HH:mm:ss"));
        cell0.setCellValue(LocalDateTime.now());
        cell0.setCellStyle(cellStyle0);

        XSSFCellStyle cellStyle1 = workbook.createCellStyle();
        cellStyle1.setDataFormat(workbook.createDataFormat().getFormat("0%"));
        cell1.setCellValue("0.5");
        cell1.setCellStyle(cellStyle1);

        OutputStream outputStream = null;

        try {
            outputStream = new FileOutputStream("/home/amanda/Documentos/novoarquivo.xlsx");
            workbook.write(outputStream);
            System.out.println("Arquivo Excel criado!");
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            workbook.close();
            outputStream.close();
        }

    }

    public static void lerColunasArqXlsx() {
        File arquivo = new File("/home/amanda/Documentos/Exercicio_Apache_Poi.xlsx");

        try (InputStream input = new FileInputStream(arquivo)) {
            XSSFWorkbook workbook = new XSSFWorkbook(input);
            XSSFSheet sheet = workbook.getSheetAt(0);

            Iterator<Row> linhas = sheet.rowIterator();

            Cell coluna = null;

            while(linhas.hasNext()) {
                coluna = linhas.next().getCell(0);
                System.out.println(coluna.toString());
            }

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
