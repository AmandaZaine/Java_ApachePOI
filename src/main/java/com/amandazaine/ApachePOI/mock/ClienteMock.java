package com.amandazaine.ApachePOI.mock;

import com.amandazaine.ApachePOI.model.Cliente;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class ClienteMock {

    public static List<Cliente> getClientesMock() {
        List<Cliente> clientes = new ArrayList<>();

        clientes.add(new Cliente(1L, "João", "Silva", "12345678900", "joao.silva@gmail.com", LocalDate.of(1990, 1, 15)));
        clientes.add(new Cliente(2L, "Maria", "Oliveira", "23456789011", "maria.oliveira@yahoo.com", LocalDate.of(1985, 5, 20)));
        clientes.add(new Cliente(3L, "Carlos", "Santos", "34567890122", "carlos.santos@outlook.com", LocalDate.of(1992, 8, 10)));
        clientes.add(new Cliente(4L, "Ana", "Souza", "45678901233", "ana.souza@gmail.com", LocalDate.of(1988, 2, 28)));
        clientes.add(new Cliente(5L, "Paulo", "Pereira", "56789012344", "paulo.pereira@hotmail.com", LocalDate.of(1995, 11, 3)));
        clientes.add(new Cliente(6L, "Juliana", "Costa", "67890123455", "juliana.costa@gmail.com", LocalDate.of(1993, 7, 12)));
        clientes.add(new Cliente(7L, "Ricardo", "Almeida", "78901234566", "ricardo.almeida@yahoo.com", LocalDate.of(1990, 4, 18)));
        clientes.add(new Cliente(8L, "Fernanda", "Lima", "89012345677", "fernanda.lima@outlook.com", LocalDate.of(1987, 6, 25)));
        clientes.add(new Cliente(9L, "Rafael", "Gomes", "90123456788", "rafael.gomes@gmail.com", LocalDate.of(1989, 3, 9)));
        clientes.add(new Cliente(10L, "Beatriz", "Martins", "01234567899", "beatriz.martins@hotmail.com", LocalDate.of(1996, 12, 30)));

        return clientes;
    }

}
