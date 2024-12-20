package com.amandazaine.ApachePOI;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@SpringBootApplication
public class ApachePoiApplication {

	public static void main(String[] args) throws IOException {
		SpringApplication.run(ApachePoiApplication.class, args);
		//PoiDemonstration.saveFile();
		//PoiDemonstration.readFile();

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

		PoiDemonstration.fillInFile(dadosPessoais);
	}

}
