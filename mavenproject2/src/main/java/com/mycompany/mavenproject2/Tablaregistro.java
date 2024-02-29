/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.mycompany.mavenproject2;

/**
 *
 * @author chris
 */
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.InputMismatchException;
import java.util.Iterator;
import java.util.Scanner;

public class Tablaregistro {
    public static void main(String[] args) {
        Scanner infu = new Scanner(System.in);
        FileInputStream lector = null;
        Workbook datos = null;
        Sheet hoja = null;
        try {
            lector = new FileInputStream("C:/Users/chris/Documents/tabla.xlsx");
            datos = new XSSFWorkbook(lector);
            hoja = datos.getSheetAt(0);
        } catch (IOException e) {
            System.out.println("Error al abrir el archivo de Excel. Asegúrate de que el archivo existe y es accesible.");
            return;
        }
        while (true) {
            System.out.println("1-Nuevo reguistro. 2-Ayudas. 3-Cursos. 4-Imprimir Reguistro. 5-Salir.");
            System.out.println("Selecciona el proceso a realizar.");
            int selc = 0;
            try {
                selc = infu.nextInt();
                infu.nextLine(); // consume newline
            } catch (InputMismatchException e) {
                System.out.println("Entrada inválida. Por favor, introduce un número.");
                infu.nextLine(); // consume the invalid input
                continue;
            }
            switch (selc) {
                case 1:
                    registrarCliente(infu, hoja);
                    break;
                // case 2, 3, 4, 5: falta las opciones 
                default:
                    System.out.println("Opción no válida.");
            }
            System.out.println("¿Deseas realizar otra operación? (SI/NO)");
            String respuesta = infu.nextLine();
            if (!respuesta.equalsIgnoreCase("SI")) {
                break;
            }
        }
        try {
            FileOutputStream escribir = new FileOutputStream("C:/Users/chris/Documents/tabla.xlsx");
            datos.write(escribir);
            datos.close();
        } catch (IOException e) {
            System.out.println("Error al guardar el archivo de Excel. Asegúrate de que el archivo no está abierto en otro programa.");
        }
    }
    private static void registrarCliente(Scanner infu, Sheet hoja) {
        System.out.println("Ingresa el nombre del cliente: ");
        String nombre = infu.nextLine().toUpperCase();
        System.out.println("Ingresa edad del cliente: ");
        int edad = 0;
        try { 
            edad = infu.nextInt();
            infu.nextLine(); // consume newline
        } catch (InputMismatchException e) {
            System.out.println("Entrada inválida. Por favor, introduce un número.");
            infu.nextLine(); // consume the invalid input
            return;
        }
           if (edad<18){
        System.out.println("Menor de edad no se permite registro.");
        return;
        }
        System.out.println("Ingresa la direccion: ");
        String direccion = infu.nextLine().toUpperCase();
        System.out.println("Ingresa la Ayuda para el cliente: ");
        String ayuda = infu.nextLine().toUpperCase();

        Iterator<Row> iterador = hoja.iterator();
        while (iterador.hasNext()) {
            Row sigfila = iterador.next();
            Cell celnombre = sigfila.getCell(0);
            if (celnombre != null && celnombre.getStringCellValue().equals(nombre)) {
                System.out.println("El cliente ya está registrado. Aquí están los detalles:");
                System.out.println("Nombre: " + celnombre.getStringCellValue());
                System.out.println("Edad: " + sigfila.getCell(1).getNumericCellValue());
                System.out.println("Direccion: " + sigfila.getCell(2).getStringCellValue());
                System.out.println("Ayuda: " + sigfila.getCell(7).getStringCellValue());
                return;
            }
        }
        int contador = hoja.getLastRowNum();
        Row fila = hoja.createRow(++contador);
        fila.createCell(0).setCellValue(nombre);
        fila.createCell(1).setCellValue(edad);
        fila.createCell(7).setCellValue(ayuda);
        fila.createCell(2).setCellValue(direccion);

        System.out.println("Cliente registrado exitosamente");
    }
}