package com;
/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */



 import org.apache.poi.ddf.EscherColorRef.SysIndexProcedure;
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
 import javax.imageio.ImageIO;
 import org.apache.commons.io.IOUtils;
 import javax.swing.*;
 import java.awt.*;
import java.io.ByteArrayInputStream;
import java.util.List;
import java.awt.BorderLayout;
import java.awt.image.BufferedImage;
import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStream;
import java.awt.Graphics;
import javax.imageio.ImageIO;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.ImageIcon;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFShape;
 
 public class tablaregistro1  {
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
         // Añade una lista para almacenar todas las ventanas de imágenes
        List<JFrame> frames = new ArrayList<>();
        Scanner infu = new Scanner(System.in);
         FileInputStream lector = null;
         Workbook datos = null;
         Sheet hoja = null;
         try {
             lector = new FileInputStream("C:/Users/chris/OneDrive/Documents/Libro2.xlsx");
             datos = new XSSFWorkbook(lector);
             hoja = datos.getSheetAt(0);
         } catch (IOException e) {
             System.out.println(rojo+"Error al abrir el archivo de Excel. Asegúrate de que el archivo existe y es accesible."+reset);
             return;
         }
         while (true) {
             System.out.print("\033[H\033[2J"); 
             System.out.flush();
             System.out.println(amarillo+"1-Nuevo reguistro. 2-Nueva busqueda. 3-Eliminar elemento. 4-Salir."+reset);
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
                     registrarCliente(infu, hoja,columna,datos);
                     break;
                 case 2: 
                 int columnamar=0, columnamod=1;
                 String marca = leer(infu, "Ingresa Marca: ");
                 String modelo = camposobligatorios(infu, "Ingresa Modelo: ");
                    buscarauto(marca, modelo, hoja,columnamar, columnamod,frames);
                 
                 break;
                 case 3:
                 int reg;
                 do {
                    System.out.println("1-Registro General.   2-Registro Individual."); 
                    while (!infu.hasNextInt()) {
                        System.out.println(rojo+"Por favor, selecciona una opcion."+reset);
                    infu.next(); // descarta la entrada incorrecta
                }
                reg=infu.nextInt();
                infu.nextLine(); // consume newline
            } while (reg <= 0 || reg>2);
                 switch (reg) {
                    case 1:
                        for(Row fila:hoja){
                            for(Cell celdas:fila){
                                String valores="";
                                switch (celdas.getCellType()) {
                                    case STRING:
                                        valores=celdas.getStringCellValue();
                                        break;
                                case NUMERIC:
                                valores=String.valueOf(celdas.getNumericCellValue());
                                break;
                                    default:
                                        break;
                                }
                                System.out.println(valores);
                            }
                            System.out.println();
                        }
                        
                        break;
                        case 2:
                       marca = leer(infu, "Ingresa el nombre del cliente a buscar: ");
                        buscar(marca,hoja,0);
                        break;
                    default:
                        break;
                 }
                 break;
                 case 4:
                 break; 
                 default:
                     System.out.println(rojo+"Opción no válida."+reset);
             }
           
             System.out.println("¿Deseas realizar otra operación? (SI/NO)");
             String respuesta = infu.nextLine();
             if (!respuesta.equalsIgnoreCase("SI")) {
                 // Cierra todas las ventanas de imágenes
        for (JFrame marco : frames) {
            marco.dispose();
        }
        break;

             }
         }
         guardarInformacion(datos, "C:/Users/chris/OneDrive/Documents/Libro2.xlsx");
         cerrarArchivo(datos);
     }
     public static void cerrarArchivo(Workbook datos) {
        try {
            datos.close();
        } catch (IOException e) {
            System.out.println(rojo + "Error al cerrar el archivo de Excel." + reset);
        }
    }    
    public static void guardarInformacion(Workbook datos, String rutaArchivo) {
        try {
            FileOutputStream escribir = new FileOutputStream(rutaArchivo);
            datos.write(escribir);
            escribir.close(); // Es importante cerrar el FileOutputStream después de usarlo
        } catch (IOException e) {
            System.out.println(rojo + "Error al guardar el archivo de Excel. Asegúrate de que el archivo no está abierto en otro programa." + reset);
        }
    }
     private static void registrarCliente(Scanner infu, Sheet hoja,int columna,Workbook datos) {
 
         String marca = leer(infu, "Ingresa la Marca: ");
         String modelo  = camposobligatorios(infu, "Ingresa Modelo: ");
         String descripcion = camposobligatorios(infu, "Ingresa la Descripcion en el siguiente formato=> cilindrage, c/ de furza. ");
        //invocacion del metodo buscar para ver si el cliente ya esta registrado mandamos el nombre a buacar ynla hoja donde  debe buscar
         buscar(modelo,hoja,columna);

         int contador = hoja.getLastRowNum();
         Row fila = hoja.createRow(++contador);
         fila.createCell(0).setCellValue(marca);
         fila.createCell(1).setCellValue(modelo);
         fila.createCell(2).setCellValue(descripcion);
         
         
         System.out.println(verde+"Registro exitoso"+reset);
         guardarInformacion(datos, "C:/Users/chris/OneDrive/Documents/Libro2.xlsx");
     }
     public static void buscarauto(String marca, String modelo, Sheet hoja, int columnamarc, int columnamod, List<JFrame> frames ) {
        Iterator<Row> iterador = hoja.iterator();
        while (iterador.hasNext()) {
            Row sigfila = iterador.next();
            Cell celmarca = sigfila.getCell(columnamarc);
            Cell celmodelo = sigfila.getCell(columnamod);
            if (celmarca != null && celmarca.getStringCellValue().equalsIgnoreCase(marca) && celmodelo != null && celmodelo.getStringCellValue().equalsIgnoreCase(modelo)) {
                System.out.println("Registrado existente. Aquí están los detalles:");
                System.out.println("Marca: " + celmarca.getStringCellValue());
                if (sigfila.getCell(1).getCellType() == CellType.STRING) {
                    System.out.println("Modelo: " + sigfila.getCell(1).getStringCellValue());
                } else if (sigfila.getCell(1).getCellType() == CellType.NUMERIC) {
                    System.out.println("Modelo: " + sigfila.getCell(1).getNumericCellValue());
                }
                System.out.println("Descripcion: " + sigfila.getCell(2).getStringCellValue());
            buscarimagen(hoja,sigfila,frames);
            }
        }
    }
    public static void buscarimagen(Sheet hoja, Row sigfila, List<JFrame> frames ){

        // Buscar la imagen en la columna E
                if (hoja instanceof XSSFSheet) {
                    XSSFSheet cashoja = (XSSFSheet) hoja;
                    for (XSSFShape imagenes : cashoja.createDrawingPatriarch().getShapes()) {
                        if (imagenes instanceof XSSFPicture) {
                            XSSFPicture imagen = (XSSFPicture) imagenes;
                            XSSFClientAnchor anchor = imagen.getPreferredSize();
                            int filaimagen1 = anchor.getRow1();
                            int columnaimagen1 = anchor.getCol1();
                            if (filaimagen1 == sigfila.getRowNum() && columnaimagen1 == 4) { // Columna E
                                byte[] datosimagen = imagen.getPictureData().getData();
                            /*crear un panel para mostrar la imagen
                            JPanel panel=new JPanel(){
                                private BufferedImage image;
                                public void setimagen(BufferedImage image){
                                    this.image=image;
                                    repaint();
                                }

                            }*/
                                // Convierte los bytes de la imagen en una imagen de BufferedImage
                                try {
                                    InputStream is = new ByteArrayInputStream(datosimagen);
                                    BufferedImage image = ImageIO.read(is);
    
                                    // Muestra la imagen en un JFrame
                                    JFrame marco = new JFrame();
                                    marco.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);
                                    JLabel etiqueta = new JLabel(new ImageIcon(image));
                                    marco.getContentPane().add(etiqueta, BorderLayout.CENTER);
                                    marco.pack();
                                    marco.setVisible(true);
                                    // Añade el marco a la lista de marcos
                                    frames.add(marco);

                                } catch (IOException e) {
                                    e.printStackTrace();
                                }
                            }
                        }
                    }
                }
            }
        
    
     public static void buscar(String nombre,Sheet hoja, int columna) {
        
         Iterator<Row> iterador = hoja.iterator();
         while (iterador.hasNext()) {
             Row sigfila = iterador.next();
             Cell celnombre = sigfila.getCell(columna);
             if (celnombre != null && celnombre.getStringCellValue().equalsIgnoreCase(nombre)) {
                 System.out.println(azul+"Registrado existente. Aquí están los detalles:");
                 System.out.println("Marca: " + celnombre.getStringCellValue());
                 System.out.println("Modelo: " + sigfila.getCell(1).getNumericCellValue());
                 System.out.println("Descripcion: " + sigfila.getCell(2).getStringCellValue()+reset); 
             }
         }
     }

     private static String camposobligatorios(Scanner infu, String mensaje) {
         String dato;//ya sea el nombre o la direccion
         do {
             System.out.println(mensaje);
             dato = infu.nextLine().toUpperCase().trim();
             if (dato.isEmpty()) {
                 System.out.println(rojo+"Campo Obligatorio."+reset);
             }
         } while (dato.isEmpty());
         return dato;
     }
     private static String leer(Scanner infu, String mensaje) {
      
         String dato="";//nombre
         do {
             System.out.println(mensaje);
             try{
                 dato = infu.nextLine().toUpperCase().trim();
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
     public static void emprendedores(Scanner infu, String nombre, Sheet hoja, Workbook datos) {
        int selec=0,tipocurso=0;
        String hora="", material="",descripcioncur="";
        for(Row fila:hoja){
            Cell celdanombre=fila.getCell(0);
            if(celdanombre.getCellType()==CellType.STRING&&celdanombre.getStringCellValue().equalsIgnoreCase(nombre)){    
        buscar(nombre, hoja, 0);
                do {
            System.out.println("Selecciona la ayuda para el cliente: 1-Cursos.  2-Vivienda.   3-Despensa.");
            while (!infu.hasNextInt()) {
                System.out.println(rojo+"Campo Obligatorio. Por favor, selecciona una opcion."+reset);
                infu.next(); // descarta la entrada incorrecta
            }
            selec= infu.nextInt();
            infu.nextLine(); // consume newline
        } while (selec <= 0 || selec>3);
            switch (selec) {
                case 1:
                Cell celdacomparar=fila.getCell(4);
                Cell celdadestino=fila.getCell(3);  
                do {
                    System.out.println("Descripcion del cuerso: 1-Cultora de belleza. 2-Decoracion de globos. 3-Corte y Confeccion."); 
                    while (!infu.hasNextInt()) {
                        System.out.println(rojo+"Campo Obligatorio. Por favor, selecciona una opcion."+reset);
                    infu.next(); // descarta la entrada incorrecta
                }
                tipocurso=infu.nextInt();
                infu.nextLine(); // consume newline
            } while (tipocurso <= 0 || tipocurso>3);
                System.out.println(amarillo+"Horario de Curso."+reset);
                hora=leerHoras(infu);
                switch (tipocurso) {
                    case 1:
                       descripcioncur ="CULTORA DE BELLEZA";
                        break;
                        case 2:
                        descripcioncur ="DECORACION DE GLOBOS";
                         break;
                         case 3:
                       descripcioncur ="CORTE Y CONFECCION";
                        break;
                    default:
                        break;
                }
                if(celdacomparar == null || celdacomparar.getCellType()==CellType.BLANK || celdadestino == null || celdadestino.getCellType()==CellType.BLANK){
                fila.createCell(3).setCellValue(descripcioncur);
                fila.createCell(4).setCellValue(hora);
                fila.createCell(7).setCellValue("CURSO");
                buscar(nombre, hoja, 0);
                }else{
                   
                    if(celdacomparar.getStringCellValue().equalsIgnoreCase(hora)){
                        System.out.println(rojo+"La hora del 2do curso no puede ser igual a la del 1er curso."+reset);
                    }else {
                       
                        fila.createCell(6).setCellValue(hora);
                        fila.createCell(5).setCellValue(descripcioncur);
                        fila.createCell(7).setCellValue("CURSO");
                        buscar(nombre, hoja, 0);
                    }
                }
                guardarInformacion(datos, "C:/Users/chris/OneDrive/Documents/Libro2.xlsx");
                break;
            case 2:
                
                int vivienda=0;
                String descripcionvivi="";
                
                String direccion = camposobligatorios(infu, "Ingresa la Direccion: ");
                for(Row filas:hoja){
                    Cell celdadireccion=filas.getCell(2);
                    if (celdadireccion!= null&&celdadireccion.getCellType()==CellType.STRING&&celdadireccion.getStringCellValue().equalsIgnoreCase(direccion)){
                        Cell celdaedad=filas.getCell(1);
                        if(celdaedad!=null&&celdaedad.getCellType()==CellType.NUMERIC){
                            System.out.println((int)celdaedad.getNumericCellValue());
                        for(Cell celdas:filas){
                            if(celdas.getColumnIndex()!=1){
                                if(celdas.getCellType()==CellType.NUMERIC){
                                    System.out.println((int) celdas.getNumericCellValue());
                                }else if(celdas.getCellType()==CellType.STRING){
                                    System.out.println(celdas.getStringCellValue());
                                }
                            }
                        }
                        System.out.println();
                    }
                }
            }
                for(Row filavivi:hoja){
             Cell celdireccion =filavivi.getCell(2);
             if(celdireccion.getCellType()==CellType.STRING&&celdireccion.getStringCellValue().equalsIgnoreCase(direccion)){    
                celdacomparar=filavivi.getCell(8);
                celdadestino=filavivi.getCell(9); 
             if(celdacomparar == null || celdacomparar.getCellType()==CellType.BLANK || celdadestino == null || celdadestino.getCellType()==CellType.BLANK){
                   
                        do {
                            System.out.println("Descripcion de la ayuda a vivienda: 1-Tinaco. 2-Calentador solar. 3-Material de construccion."); 
                            while (!infu.hasNextInt()) {
                                System.out.println(rojo+"Campo Obligatorio. Por favor, selecciona una opcion."+reset);
                            infu.next(); // descarta la entrada incorrecta
                        }
                        vivienda=infu.nextInt();
                        infu.nextLine(); // consume newline
                    } while (vivienda <= 0 || vivienda>3);

                    switch (vivienda) {
                        case 1:
                           descripcionvivi="TINACO";
                            break;
                            case 2:
                            descripcionvivi ="CALENTADOR SOLAR";
                             break;
                             case 3:
                           descripcionvivi ="MATERIAL DE CONSTRUCCION";
                           material=camposobligatorios(infu, "Materiales con los cuuales se apoyara al cliente.");
                            break;
                        default:
                            break;
                    }
                }
                } 
            }  
          
            guardarInformacion(datos, "C:/Users/chris/OneDrive/Documents/Libro2.xlsx");
            break;
            case 3:
            String despensa="DESPENSA";
            direccion = camposobligatorios(infu, "Ingresa la Direccion: ");
                for(Row filas:hoja){
                    Cell celdadireccion=filas.getCell(2);
                    if (celdadireccion!= null&&celdadireccion.getCellType()==CellType.STRING&&celdadireccion.getStringCellValue().equalsIgnoreCase(direccion)){
                        Cell celdaedad=filas.getCell(1);
                        if(celdaedad!=null&&celdaedad.getCellType()==CellType.NUMERIC){
                            System.out.println((int)celdaedad.getNumericCellValue());
                        for(Cell celdas:filas){
                            if(celdas.getColumnIndex()!=1){
                                if(celdas.getCellType()==CellType.NUMERIC){
                                    System.out.println((int) celdas.getNumericCellValue());
                                }else if(celdas.getCellType()==CellType.STRING){
                                    System.out.println(celdas.getStringCellValue());
                                }
                            }
                        }
                        System.out.println();
                    }
                }
            }
            for(Row filavivi:hoja){
                Cell celdireccion =filavivi.getCell(2);
                if(celdireccion.getCellType()==CellType.STRING&&celdireccion.getStringCellValue().equalsIgnoreCase(direccion)){    
                   celdacomparar=filavivi.getCell(10); 
                if(celdacomparar == null || celdacomparar.getCellType()==CellType.BLANK){
                    filavivi.createCell(10).setCellValue(despensa);
                }else{
                    System.out.println("Ya a recibido esta ayuda");
                    break;
                } 
            }
        }
        for(Row desfila:hoja){
            Cell celda=desfila.getCell(2);
            if(celda!=null&&celda.getCellType()==CellType.STRING&&celda.getStringCellValue().equalsIgnoreCase(direccion)){
                Cell descrip= desfila.createCell(10);
                descrip.setCellValue(despensa);
            }

        }
        System.out.println("Proceso de escaneo relizado.");
        
            break;
                default:
                    break;
            }
            break;
        }        
     }
     guardarInformacion(datos, "C:/Users/chris/OneDrive/Documents/Libro1.xlsx");
    }
    public static void replicar(String direccion, String descripcion, int buscar,  int guardar, Sheet hoja){
        for(Row fila:hoja){
            Cell celda=fila.getCell(buscar);
            if(celda!=null&&celda.getCellType()==CellType.STRING&&celda.getStringCellValue().equalsIgnoreCase(direccion)){
                Cell descrip= fila.createCell(guardar);
                descrip.setCellValue(descripcion);
            }

        }
        System.out.println("Proceso de escaneo relizado.");
    }
 
 }