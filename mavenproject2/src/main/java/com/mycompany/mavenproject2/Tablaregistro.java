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

import javax.print.DocFlavor.STRING;

import java.time.LocalTime;
import java.time.format.DateTimeParseException;

public class Tablaregistro {
    public static final String reset = "\u001B[0m";
    public static final String negro = "\u001B[30m";
    public static final String rojo = "\u001B[31m";
    public static final String verde= "\u001B[32m";
    public static final String amarillo = "\u001B[33m";
    public static final String azul = "\u001B[34m";
    public static final String morado = "\u001B[35m";
    public static final String azulclaro = "\u001B[36m";
    public static final String blanco = "\u001B[37m";
    public static void main(String[] args) {
        Scanner infu = new Scanner(System.in);
        FileInputStream lector = null;
        Workbook datos = null;
        Sheet hoja = null;
        try {
            lector = new FileInputStream("C:/Users/chris/OneDrive/Documents/Libro1.xlsx");
            datos = new XSSFWorkbook(lector);
            hoja = datos.getSheetAt(0);
        } catch (IOException e) {
            System.out.println(rojo+"Error al abrir el archivo de Excel. Asegúrate de que el archivo existe y es accesible."+reset);
            return;
        }
        while (true) {
            System.out.print("\033[H\033[2J"); 
            System.out.flush();
            System.out.println(amarillo+"1-Nuevo reguistro. 2-Nuevo Cursos. 3-Nueva Ayuda. 4-Imprimir Reguistro. 5-Salir."+reset);
            System.out.println();
            System.out.println("Selecciona el proceso a realizar.");
            int selc = 0;
            try {
                selc = infu.nextInt();
                infu.nextLine(); // consume newline
            } catch (InputMismatchException e) {
                System.out.println(rojo+"Entrada inválida. Por favor, introduce un número."+reset);
                infu.nextLine(); // consume the invalid input
                continue;
            }
            switch (selc) {
                case 1:
                int columna=0;
                    registrarCliente(infu, hoja,columna);
                    break;
                case 2:
                 int columnanom=0, columnacomp=4, columnacurso=5, columnahor=6, columnades=3;
                 String descripcion="",hora="";

                String nombre = camposobligatorios(infu, "Ingresa el nombre del cliente: ");
                buscar(nombre,hoja,columnanom);
               
                System.out.println("Descripcion del cuerso: Cultora de belleza, Decoracion de globos, Corte y Confeccion"); 
                descripcion=infu.nextLine().toUpperCase();
                System.out.println(amarillo+"Horario de Curso."+reset);
                hora=leerHoras(infu);
                celdas(nombre,columnanom,columnacomp,columnacurso,columnahor,columnades,hoja,hora,descripcion);
                
                break;
                // case 2, 3, 4, 5: falta las opciones 
                default:
                    System.out.println(rojo+"Opción no válida."+reset);
            }
          
            System.out.println("¿Deseas realizar otra operación? (SI/NO)");
            String respuesta = infu.nextLine();
            if (!respuesta.equalsIgnoreCase("SI")) {
                break;
            }
        }
        try {
            FileOutputStream escribir = new FileOutputStream("C:/Users/chris/OneDrive/Documents/Libro1.xlsx");
            datos.write(escribir);
            datos.close();
        } catch (IOException e) {
            System.out.println(rojo+"Error al guardar el archivo de Excel. Asegúrate de que el archivo no está abierto en otro programa."+reset);
        }
    }
    private static void registrarCliente(Scanner infu, Sheet hoja,int columna) {

        String nombre = leer(infu, "Ingresa el nombre del cliente: ");
    int edad = leerEdad(infu);
    if (edad < 18) {
        System.out.println(morado+"Menor de edad no se permite registro."+reset);
        return;
    }
    String direccion = camposobligatorios(infu, "Ingresa la Direccion: ");
    String ayuda = leer(infu, "Ingresa la Ayuda para el cliente:(curso, vivenda)  ");
    String descripcion = "",hora = "",material="", descripcionvi="";
       if(ayuda.equalsIgnoreCase("CURSO")){
System.out.println("Descripcion del cuerso: Cultora de belleza, Decoracion de globos, Corte y Confeccion"); 
descripcion=infu.nextLine().toUpperCase();
System.out.println(amarillo+"Horario de Curso."+reset);
hora=leerHoras(infu);
}
else if (ayuda.equalsIgnoreCase("VIVIENDA")){
    System.out.println("Descripcion de Ayuda para vivienda: Tinaco, Calentador Solar, Material de Construccion"); 
    descripcionvi=infu.nextLine().toUpperCase().trim();
    System.out.println();
   if(descripcion.equalsIgnoreCase("Material de Construccion")){
    System.out.println("Ingresa la lista del material con el que se apoyara a la vivienda");
    material=infu.nextLine();
   }
    }
 buscar(nombre,hoja,columna);//invocacion del metodo buscar para ver si el cliente ya esta registrado mandamos el nombre a buacar ynla hoja donde  debe buscar
        
        int contador = hoja.getLastRowNum();
        Row fila = hoja.createRow(++contador);
        fila.createCell(0).setCellValue(nombre);
        fila.createCell(1).setCellValue(edad);
        fila.createCell(7).setCellValue(ayuda);
        fila.createCell(2).setCellValue(direccion);
        fila.createCell(3).setCellValue(descripcion);
        fila.createCell(4).setCellValue(hora);
        fila.createCell(9).setCellValue(material);
        fila.createCell(8).setCellValue(descripcionvi);
        
        System.out.println(verde+"Cliente registrado exitosamente"+reset);
    }
    public static void buscar(String nombre,Sheet hoja, int columna) {
       
        Iterator<Row> iterador = hoja.iterator();
        while (iterador.hasNext()) {
            Row sigfila = iterador.next();
            Cell celnombre = sigfila.getCell(columna);
            if (celnombre != null && celnombre.getStringCellValue().equalsIgnoreCase(nombre)) {
                System.out.println(azul+"El cliente ya está registrado. Aquí están los detalles:");
                System.out.println("Nombre: " + celnombre.getStringCellValue());
                System.out.println("Edad: " + sigfila.getCell(1).getNumericCellValue());
                System.out.println("Direccion: " + sigfila.getCell(2).getStringCellValue());
                System.out.println("Ayuda: " + sigfila.getCell(7).getStringCellValue());
                System.out.print("Curso: " + sigfila.getCell(3).getStringCellValue()+"  ");
                System.out.println("Horario: " + sigfila.getCell(4).getStringCellValue()+reset);
                
                return;
    }
}
    }
    private static String camposobligatorios(Scanner infu, String mensaje) {
       ;
        String dato;//ya sea el nombre o la direccion
        do {
            System.out.println(mensaje);
            dato = infu.nextLine().toUpperCase();
            if (dato.isEmpty()) {
                System.out.println(rojo+"Campo Obligatorio."+reset);
            }
        } while (dato.isEmpty());
        return dato;
    }
    private static String leer(Scanner infu, String mensaje) {
     
        String dato="";//ya sea el nombre o la direccion
        do {
            System.out.println(mensaje);
            try{
                dato = infu.nextLine().toUpperCase();
                if (!dato.matches("[A-Za-z ]+")) { // si la entrada contiene algo más que letras y espacios
                    throw new IllegalArgumentException(rojo+"Por favor, introduce solo texto."+reset);
                }
            } catch (IllegalArgumentException e) {
                System.out.println(rojo + e.getMessage() + reset);
                dato = ""; // resetea 'dato' para que el bucle continue
            }
            if (dato.isEmpty()) {
                System.out.println(rojo+"Campo Obligatorio."+reset);
            }
        } while (dato.isEmpty());
        return dato;
    }
    
    private static int leerEdad(Scanner infu) { 
    
        int edad = 0;
        do {
            System.out.println("Ingresa edad del cliente: ");
            while (!infu.hasNextInt()) {
                System.out.println(rojo+"Campo Obligatorio. Por favor, introduce un número."+reset);
                infu.next(); // descarta la entrada incorrecta
            }
            edad = infu.nextInt();
            infu.nextLine(); // consume newline
        } while (edad <= 0);
        return edad;
    }
    private static LocalTime leerhora(Scanner infu){

        System.out.println(amarillo+"(formato HH:mm):"+reset);
        LocalTime hora = null;
        while (hora == null) {
            try {
                hora = LocalTime.parse(infu.nextLine());
            } catch (DateTimeParseException e) {
                System.out.println(rojo+"Hora inválida. Por favor, introduce una hora válida (formato HH:mm):"+reset);
            }
        }
        return hora;
    }
    private static String leerHoras(Scanner infu) {
        LocalTime horaInicio = null, horaFin = null;
        
        while (horaInicio == null) {
            System.out.println(amarillo+"Por favor, introduce una hora de inicio:"+reset);
            
            horaInicio = leerhora(infu);
        }
        while (horaFin == null) {
            System.out.println(amarillo+"Por favor, introduce una hora de fin:"+reset);
            
            horaFin = leerhora(infu);
        }
        return horaInicio + "-" + horaFin;
    }
    public static void celdas(String nombre,int columnanom, int columnacomp,int columnacurso,int columnahor, int columnades, Sheet hoja,String dato,String descripcion){
        for(Row fila:hoja){
            Cell celda=fila.getCell(columnanom);
            if(celda.getCellType()==CellType.STRING&&celda.getStringCellValue().equalsIgnoreCase(nombre)){
                Cell celdacomp=fila.getCell(columnacomp);
                Cell celdades=fila.getCell(columnades);
                if(celdacomp == null || celdacomp.getCellType()==CellType.BLANK || celdades == null || celdades.getCellType()==CellType.BLANK){
                    if(celdacomp == null){
                        celdacomp = fila.createCell(columnacomp);
                    }
                    celdacomp.setCellValue(dato);
                    if(celdades == null){
                        celdades = fila.createCell(columnades);
                    }
                    celdades.setCellValue(descripcion);
                }else{
                    if(celdacomp.getStringCellValue().equalsIgnoreCase(dato)){
                        System.out.println("La hora del 2do curso no puede ser igual a la del 1er curso.");
                        System.out.println("El domicilio ya se encuentra registrado");
                    }else {
                        Cell celdaguar=fila.createCell(columnahor);
                        celdaguar.setCellValue(dato);
                        Cell celdaguarc=fila.createCell(columnacurso);
                        celdaguarc.setCellValue(descripcion);
                    }
                }
                break;
            }
        }
    }



}