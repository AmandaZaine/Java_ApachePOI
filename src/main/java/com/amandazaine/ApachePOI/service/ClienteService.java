package com.amandazaine.ApachePOI.service;

import com.amandazaine.ApachePOI.model.Cliente;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.util.List;

public class ClienteService {

    public void relatorioClientes(List<Cliente> clientes) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Clientes");
        XSSFRow row = null;
        Field[] headers = Cliente.class.getDeclaredFields();

        for(int i = 0; i < clientes.size(); i++) {
            if(i == 0) {
                row = sheet.createRow(i);

                for(int j = 0; j < headers.length; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(headers[j].getName());
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
                    XSSFCellStyle dateCellStyle = workbook.createCellStyle();
                    dateCellStyle.setDataFormat(workbook.createDataFormat().getFormat("dd/MM/yyyy"));

                    cell.setCellValue((LocalDate) atributosCliente.get(k));
                    cell.setCellStyle(dateCellStyle);
                }

                sheet.autoSizeColumn(k);
            }
        }

        try {
            OutputStream outputStream = new FileOutputStream("/home/amanda/Documentos/Exercicio_Apache_Poi.xlsx");
            workbook.write(outputStream);

            workbook.close();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("Erro ao criar documento: " + e);
        }

    }
}
