package com.amandazaine.ApachePOI.service;

import com.amandazaine.ApachePOI.designPatterns.ApachePOIBuilder;
import com.amandazaine.ApachePOI.model.Cliente;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.util.List;

@Service
public class ClienteService {

    public void gerarRelatorioClientes(List<Cliente> clientes) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Clientes");

        XSSFFont fonteHeader = new ApachePOIBuilder.FontBuilder()
                               .setNegrito(true)
                               .build(workbook);

        XSSFCellStyle estiloHeader = new ApachePOIBuilder.CellStyleBuilder()
                                    .setCorPadrao(IndexedColors.AQUA.getIndex())
                                    .setTipoPreenchimento(FillPatternType.SOLID_FOREGROUND)
                                    .setAlinhamentoHorizontal(HorizontalAlignment.CENTER)
                                    .setEstiloBordaSuperior(BorderStyle.MEDIUM)
                                    .setEstiloBordaInferior(BorderStyle.MEDIUM)
                                    .setEstiloBordaEsquerda(BorderStyle.MEDIUM)
                                    .setEstiloBordaDireita(BorderStyle.MEDIUM)
                                    .setFonte(fonteHeader)
                                    .build(workbook);

        XSSFCellStyle estiloData = new ApachePOIBuilder.CellStyleBuilder().setFormatoDadoCelula("dd/MM/yyyy").build(workbook);

        XSSFRow row = null;
        Field[] headers = Cliente.class.getDeclaredFields();

        for(int i = 0; i < clientes.size(); i++) {
            if(i == 0) {
                row = sheet.createRow(i);

                for(int j = 0; j < headers.length; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(headers[j].getName());
                    cell.setCellStyle(estiloHeader);
                }
            }

            Cliente cliente = clientes.get(i);
            List<Object> atributosCliente = cliente.obterAtributos();

            row = sheet.createRow(i + 1);

            for (int k = 0; k < atributosCliente.size(); k++) {
                Cell cell = row.createCell(k);

                if (atributosCliente.get(k) instanceof Long) {
                    cell.setCellValue((Long) atributosCliente.get(k));
                }
                if (atributosCliente.get(k) instanceof String) {
                    cell.setCellValue((String) atributosCliente.get(k));
                }
                if (atributosCliente.get(k) instanceof LocalDate) {
                    cell.setCellValue((LocalDate) atributosCliente.get(k));
                    cell.setCellStyle(estiloData);
                }

                sheet.autoSizeColumn(k);
            }
        }

        try {
            OutputStream outputStream = new FileOutputStream("/home/amanda/Documentos/Exercicio_Apache_Poi_.xlsx");
            workbook.write(outputStream);

            workbook.close();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("Erro ao criar documento: " + e);
        }
    }
}
