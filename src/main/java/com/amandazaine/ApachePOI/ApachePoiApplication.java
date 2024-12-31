package com.amandazaine.ApachePOI;

import com.amandazaine.ApachePOI.model.Cliente;
import com.amandazaine.ApachePOI.service.ClienteService;
import org.apache.commons.codec.DecoderException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static com.amandazaine.ApachePOI.mock.ClienteMock.getClientesMock;

//@SpringBootApplication
public class ApachePoiApplication {

	public static void main(String[] args) throws IOException, ParseException, DecoderException {
		//SpringApplication.run(ApachePoiApplication.class, args);

		List<Cliente> clientes = getClientesMock();

		ClienteService clienteService = new ClienteService();
		clienteService.relatorioClientes(clientes);

	}

	private static List<List<String>> getListsOfDate() {
		List<List<String>> datas = Arrays.asList(
				Arrays.asList("25/12/2024", "11/12/2025"),
				Arrays.asList("01/01/2025", "17/08/2030"),
				Arrays.asList("14/02/2025", "05/07/2027")
		);
		return datas;
	}

	private static List<List<String>> listaDadosPessoais() {
		List<List<String>> dadosPessoais = new ArrayList<>();

		List<String> registro1 = new ArrayList<>();
		registro1.add("João Silva");
		registro1.add("123.456.789-00");
		registro1.add("São Paulo");

		List<String> registro2 = new ArrayList<>();
		registro2.add("Maria Oliveira");
		registro2.add("987.654.321-11");
		registro2.add("Rio de Janeiro");

		List<String> registro3 = new ArrayList<>();
		registro3.add("Carlos Mendes");
		registro3.add("456.789.123-22");
		registro3.add("Belo Horizonte");

		dadosPessoais.add(registro1);
		dadosPessoais.add(registro2);
		dadosPessoais.add(registro3);
		return dadosPessoais;
	}

}
