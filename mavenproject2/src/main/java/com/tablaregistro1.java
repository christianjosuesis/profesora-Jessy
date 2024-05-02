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
import org.apache.xmlbeans.XmlCursor;

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
 import java.awt.event.*;
 import java.awt.*;
import java.io.ByteArrayInputStream;
import java.util.List;
import java.awt.image.BufferedImage;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStream;
import java.awt.Color;
import java.awt.Font;

import javax.imageio.ImageIO;

import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFShape;
 
 public class tablaregistro1  extends JPanel  {
    private BufferedImage[] imagenes = new BufferedImage[3];
    private int imagenActual = 0; // Esta variable mantendrá la posición actual de la imagen
    private JButton boton1;
    private JButton boton2;
    private JButton boton3;
    private String marca;
    private int ano;
    private String modelot;
    private String descripcion;

    public tablaregistro1(BufferedImage imagen,BufferedImage imagen1, BufferedImage imagen2,  String marca, String modelot, int ano,String descripcion) {
        this.imagenes[0] = imagen;
        this.imagenes[1] = imagen1;
        this.imagenes[2] = imagen2;
        this.marca = marca;
        this.ano = ano;
        this.modelot = modelot;
        this.descripcion = descripcion;
         //Configurar el layout del panel principal
         this.setLayout(null);
        // Crear los botones
        boton1 = new JButton("Azul.");
        boton2 = new JButton("Rojo.");
        boton3 = new JButton("Negro.");
           
        // Crear un nuevo panel para los botones
        JPanel panelBotones = new JPanel();

        // Agregar los botones al panel de botones
        panelBotones.add(boton1);
        panelBotones.add(boton2);
        panelBotones.add(boton3);

        // Configurar el layout del panel principal
        this.setLayout(new BoxLayout(this, BoxLayout.Y_AXIS));

        // Agregar un Box.Filler para empujar el panel de botones hacia abajo
        this.add(new Box.Filler(new Dimension(0, 0), new Dimension(0, Integer.MAX_VALUE), new Dimension(0, Integer.MAX_VALUE)));

        // Agregar el panel de botones al panel principal
        this.add(panelBotones);

        /*  Configurar las coordenadas y el tamaño de los botones
        boton1.setBounds(100, 400, 80, 30);
        boton2.setBounds(200, 400, 80, 30);
        boton3.setBounds(300, 400, 80, 30);

        // Agregar los botones al panel
         this.add(boton1);
         this.add(boton2);
         this.add(boton3);*/
        
          // Agregar los listeners a los botones
        boton1.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Cambiar la imagen cuando se haga clic en el botón 1
                imagenActual=0;
                repaint(); // Redibujar el panel con la nueva imagen
            }
        });

        boton2.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Cambiar la imagen cuando se haga clic en el botón 2
                imagenActual=1;
                repaint(); // Redibujar el panel con la nueva imagen
            }
        });

        boton3.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Cambiar la imagen cuando se haga clic en el botón 3
                imagenActual=2;
                repaint(); // Redibujar el panel con la nueva imagen
            }
        });
    }

    protected void paintComponent(Graphics g) {
        super.paintComponent(g);

    // Obtener el ancho y el alto del componente
    int anchoComponente = this.getWidth();
    int altoComponente = this.getHeight();

    // Calcular el tamaño y la posición de la imagen
    int anchoImagen = anchoComponente-60;
    int altoImagen = altoComponente / 2;
    int xImagen = 30;
    int yImagen = 90;

    // Dibujar la imagen 
    g.drawImage(imagenes[imagenActual], xImagen, yImagen, anchoImagen, altoImagen, this);

    // Configurar el color y la fuente del texto
    g.setColor(Color.BLACK);
    g.setFont(new Font("Bernard MT Condensed", Font.BOLD, 30));

    // Calcular la posición del texto
    int xTexto = anchoComponente / 2-100;
    int yTexto = g.getFont().getSize()+20;

    // Dibujar el texto "marca" y "modelo" en la parte superior central
    g.drawString(marca, xTexto+40, yTexto);
    g.drawString("Modelo: " + ano, xTexto, yTexto + g.getFont().getSize());

    g.setColor(Color.BLACK);
    g.setFont(new Font("Britannic Bold", Font.BOLD, 24));
    // Dibujar el texto restante en la parte central derecha
    g.drawString(modelot, xTexto+40, 3* altoComponente / 4);
    g.setFont(new Font("Britannic Bold", Font.BOLD, 18));
    g.drawString(descripcion, anchoComponente / 7, 3* altoComponente / 4+30);
}

         /*  Ajustar el tamaño de la imagen (por ejemplo, al doble del tamaño original)
        int anchoImagen = this.getWidth();
        int altoImagen = this.getHeight();

        g.drawImage(imagen, 0, 0, anchoImagen, altoImagen, this);

        // Agregar texto en la imagen
        g.setColor(Color.BLACK); // Color del texto
        g.setFont(new Font("Drag Racing", Font.BOLD, 24)); // Fuente y tamaño del texto
        g.drawString(marca, 1000, 50); // Coordenadas (x, y) del texto
        g.drawString("Modelo: " + modelon, 1000, 80); // Coordenadas (x, y) del texto
        g.drawString(modelot, 1000, 100); // Coordenadas (x, y) del texto
        g.drawString(descripcion, 1000, 120);
    }*/
    
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
             lector = new FileInputStream("C:/Users/chris/OneDrive/Documents/baseautos.xlsx");
             datos = new XSSFWorkbook(lector);
             hoja = datos.getSheetAt(0);
         } catch (IOException e) {
             System.out.println(rojo+"Error al abrir el archivo de Excel. Asegúrate de que el archivo existe y es accesible."+reset);
             return;
         }
         int salida=0;
         boolean salidab=true;
         while (salidab==true) {
             System.out.print("\033[H\033[2J"); 
             System.out.flush();
             System.out.println(amarillo+"1-Nuevo reguistro. 2-Nueva busqueda. 3-Eliminar elemento. 4-Salir."+reset);
             System.out.println();
             System.out.println("Selecciona el proceso a realizar.");
             int selc = 0;
            /*  GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
        String[] fontFamilies = ge.getAvailableFontFamilyNames();
        for (String fontFamily : fontFamilies) {
            System.out.println(fontFamily);
        }*/
             try {
                 selc = infu.nextInt();
                 infu.nextLine(); // consume newline
             } catch (InputMismatchException e) {
                 System.out.println(rojo+"Entrada inválida. Por favor, introduce un número."+reset);
                 infu.nextLine(); // consume la entrada invalida
                 continue;
             }
             switch (selc) {
                 case 1:
                 int columna=1;
                     registrarCliente(infu, hoja,columna,datos);
                     break;
                 case 2: 
                 int columnamar=0, columnamod=1;
                 String marca = leer(infu, "Ingresa Marca: ");
                 String modelo = camposobligatorios(infu, "Ingresa Modelo: ");
                    buscarauto(infu,marca, modelo, hoja,columnamar, columnamod,frames,datos);
                 
                 break;
                 case 3:
                 columnamar=0;
                 columnamod=1;
                  marca = leer(infu, "Ingresa Marca: ");
                  modelo = camposobligatorios(infu, "Ingresa Modelo: ");
                  eliminar(marca,modelo,hoja,columnamar,columnamod,frames);
                 break;
                 case 4:
                  salida=1;
                  salidab=false;
                 break; 
                 default:
                     System.out.println(rojo+"Opción no válida."+reset);
             }
           if(salida != 1){
            String respuesta;
            do {
                System.out.println("¿Deseas realizar otra operación? (SI/NO)");
                respuesta = infu.nextLine().toUpperCase();
                if (!respuesta.equals("SI") && !respuesta.equals("NO")) {
                    System.out.println("Por favor, solo responde SI o NO.");
                }
            } while (!respuesta.equals("SI") && !respuesta.equals("NO"));
            
             if (!respuesta.equalsIgnoreCase("SI")) {
                 // Cierra todas las ventanas de imágenes
        for (JFrame marco : frames) {
            marco.dispose();
        }
        break;

             }

            }
           
         }
         guardarInformacion(datos, "C:/Users/chris/OneDrive/Documents/baseautos.xlsx");
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
        // Crear un Drawing patria en la hoja
        Drawing<?> drawing = hoja.createDrawingPatriarch();
         String marca = leer(infu, "Ingresa la Marca: ");
         String modelo  = camposobligatorios(infu, "Ingresa Modelo: ");
         int ano  = campo(infu, "Ingresa Año: ");
         String descripcion = camposobligatorios(infu, "Ingresa la Descripcion en el siguiente formato=> cilindrage, c/ de furza. ");
         System.out.println("Ingresa la Ruta de la primer imagen.");
         String rutaimagen1= infu.nextLine();
         System.out.println("Ingresa la Ruta de la segunda imagen.");
         String rutaimagen2=infu.nextLine();
         System.out.println("Ingresa la Ruta de la tercer imagen.");
         String rutaimagen3=infu.nextLine();
         /*String rutaimagen1 = ruta(infu, "Ingresa la Ruta de la primer imagen. ");
         String rutaimagen2 = ruta(infu, "Ingresa la Ruta de la segunda imagen. ");
         String rutaimagen3 = ruta(infu, "Ingresa la Ruta de la tercer imagen. ");*/
         //invocacion del metodo buscar para verificar la existencia de losd atos
        boolean registro= buscar(modelo,hoja,columna);
         if(registro==false){
// Obtener la última fila
int ultimafila = hoja.getLastRowNum();
Row nuevafila = hoja.getRow(ultimafila);
// Obtener el alto de la última fila
float altoultimafila = nuevafila.getHeightInPoints();

         int contador = ultimafila;
         Row fila = hoja.createRow(++contador);
            // Establecer el alto de la nueva fila al alto de la última fila
         fila.setHeightInPoints(altoultimafila);
         fila.createCell(0).setCellValue(marca);
         fila.createCell(1).setCellValue(modelo);
         fila.createCell(2).setCellValue(ano);
         fila.createCell(3).setCellValue(descripcion);

         try {
            // Crear un InputStream para cada imagen
            InputStream inputStream1 = new FileInputStream(rutaimagen1);
            InputStream inputStream2 = new FileInputStream(rutaimagen2);
            InputStream inputStream3 = new FileInputStream(rutaimagen3);
    
            // Convertir las imágenes en un array de bytes
            byte[] bytes1 = IOUtils.toByteArray(inputStream1);
            byte[] bytes2 = IOUtils.toByteArray(inputStream2);
            byte[] bytes3 = IOUtils.toByteArray(inputStream3);
    
            // Agregar las imágenes al Workbook y obtener los índices
            int pictureIdx1 = datos.addPicture(bytes1, Workbook.PICTURE_TYPE_JPEG);
            int pictureIdx2 = datos.addPicture(bytes2, Workbook.PICTURE_TYPE_JPEG);
            int pictureIdx3 = datos.addPicture(bytes3, Workbook.PICTURE_TYPE_JPEG);
    
            // Crear un ClientAnchor para cada imagen
            ClientAnchor anchor1 = drawing.createAnchor(0, 0, 0, 0, 4, contador, 5, contador + 1);
            ClientAnchor anchor2 = drawing.createAnchor(0, 0, 0, 0, 5, contador, 6, contador + 1);
            ClientAnchor anchor3 = drawing.createAnchor(0, 0, 0, 0, 6, contador, 7, contador + 1);
    
            // Crear un Picture para cada imagen y agregarlo al Drawing
            Picture picture1 = drawing.createPicture(anchor1, pictureIdx1);
            Picture picture2 = drawing.createPicture(anchor2, pictureIdx2);
            Picture picture3 = drawing.createPicture(anchor3, pictureIdx3);
           
// Establecer el ancho y alto deseados para las celdas
int celdaancho = hoja.getColumnWidth(4);
// Obtener el ancho y alto de la celda en puntos
int celdapuntos = celdaancho / 256; // 256 es la conversión de unidades de medida de Excel a puntos
float celdapuntosa = altoultimafila;

// Calcular el ancho y alto deseados para la imagen
int imagenpuntos = celdapuntos > 50 ? celdapuntos - 50 : celdapuntos;
float imagenpuntosa = celdapuntosa > 50 ? celdapuntosa - 50 : celdapuntosa;

// Convertir los puntos a la escala de Apache POI (1 punto = 1/72 pulgadas)
// Establecer el ancho y alto deseados para las imágenes
double imgenancho = imagenpuntos / 72.0;
double imgenalto = imagenpuntosa / 72.0;

            // Establecer el ancho de las columnas
            hoja.setColumnWidth(4, celdaancho); // para picture1
            hoja.setColumnWidth(5, celdaancho); // para picture2
            hoja.setColumnWidth(6, celdaancho); // para picture3
// Ajustar el tamaño de las imágenes
picture1.resize(imgenancho ,imgenalto);
picture2.resize(imgenancho, imgenalto);
picture3.resize(imgenancho, imgenalto);
    
            // Cerrar los InputStream
            inputStream1.close();
            inputStream2.close();
            inputStream3.close();
        } catch (Exception e) {
            e.printStackTrace();
        }         
         System.out.println(verde+"Registro exitoso"+reset);
         guardarInformacion(datos, "C:/Users/chris/OneDrive/Documents/baseautos.xlsx");
    }
        }
     public static void buscarauto(Scanner infu,String marca, String modelo, Sheet hoja, int columnamarc, int columnamod, List<JFrame> frames,  Workbook datos) {
        String Marca;
        String Descripcion;
        String Modelot;
        int Ano=0;
        String Anos;
        Iterator<Row> iterador = hoja.iterator();
        boolean registroExistente = false;
        while (iterador.hasNext()) {
            Row sigfila = iterador.next();
            Cell celmarca = sigfila.getCell(columnamarc);
            Cell celmodelo = sigfila.getCell(columnamod);
            if (celmarca != null && celmarca.getStringCellValue().equalsIgnoreCase(marca) && celmodelo != null && celmodelo.getStringCellValue().equalsIgnoreCase(modelo)) {
                System.out.println("Registrado existente. Aquí están los detalles:");
                System.out.println("Marca: " + sigfila.getCell(0).getStringCellValue());

                Marca=celmarca.getStringCellValue();

                System.out.println("Modelo: " + sigfila.getCell(1).getStringCellValue());
                    
                Modelot=sigfila.getCell(1).getStringCellValue();
                Cell cell = sigfila.getCell(2); // obtén la celda
                if (cell != null) {
                switch (cell.getCellType()) {
                    case STRING:
                    System.out.println("Año: " + sigfila.getCell(2).getStringCellValue());

                    Anos= sigfila.getCell(2).getStringCellValue();
                    Ano = Integer.parseInt(Anos);
                        // maneja el valor de la cadena
                        break;
                    case NUMERIC:
                    System.out.println("Año: " + sigfila.getCell(2).getNumericCellValue());

                        Ano = (int) cell.getNumericCellValue();
                        // maneja el valor numérico
                        break;
                    default:
                        System.out.println("Tipo de celda no soportado");
                }
            }else {
                System.out.println("La celda es null");
            }
                System.out.println("Descripcion: " + sigfila.getCell(3).getStringCellValue());

                Descripcion=sigfila.getCell(3).getStringCellValue();
                // Buscar la imagen
                BufferedImage image=null,image1=null,image2=null;
                 for(int i=4;i<=6;i++){
                    if(i==4){
                    image= buscarImagen(hoja, sigfila,i);
                    }else if(i==5){
                    image1 = buscarImagen(hoja, sigfila,i);
                 }else if(i==6){
                    image2 = buscarImagen(hoja, sigfila,i);
                }
            }
                 if (image != null || image1 != null || image2 != null) {
                // Crear una instancia de MiPanel
                
                tablaregistro1 panel = new tablaregistro1(image, image1,image2,Marca, Modelot,Ano, Descripcion);
                
                // Crear un JFrame para mostrar el panel
                JFrame frame = new JFrame();
                frame.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);
                frame.add(panel);
                frame.setSize(500,500);
                frame.setLocationRelativeTo(null);
                frame.setVisible(true);

                // Añadir el marco a la lista de marcos
                frames.add(frame);

                 }
             registroExistente = true;
        }
    }
        if(!registroExistente){
            
            String respuesta;
do {
    System.out.println("Registro inexistente. ¿Deseas crear un nuevo registro? (SI/NO)");
    respuesta = infu.nextLine().toUpperCase();
    if (!respuesta.equals("SI") && !respuesta.equals("NO")) {
        System.out.println("Por favor, solo responde SI o NO.");
    }
} while (!respuesta.equals("SI") && !respuesta.equals("NO"));
if(respuesta.equalsIgnoreCase("SI")){
    registrarCliente(infu, hoja,columnamod,datos);
}
        }
        }
   public static BufferedImage buscarImagen(Sheet hoja, Row sigfila,int columna) {
    // Buscar la imagen en la columna E
    if (hoja instanceof XSSFSheet) {
        XSSFSheet cashoja = (XSSFSheet) hoja;
        for (XSSFShape imagenes : cashoja.createDrawingPatriarch().getShapes()) {
            if (imagenes instanceof XSSFPicture) {
                XSSFPicture imagen = (XSSFPicture) imagenes;
                XSSFClientAnchor anchor = imagen.getPreferredSize();
                int filaimagen1 = anchor.getRow1();
                int columnaimagen1 = anchor.getCol1();
                if (filaimagen1 == sigfila.getRowNum() && columnaimagen1 == columna) { // Columna E
                    byte[] datosimagen = imagen.getPictureData().getData();

                    // Convierte los bytes de la imagen en una imagen de BufferedImage
                    try {
                        InputStream is = new ByteArrayInputStream(datosimagen);
                        BufferedImage image = ImageIO.read(is);
                        return image;
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        }
    }
    return null;
}
public static void mostrarImagen(Component image, List<JFrame> frames) {
   SwingUtilities.invokeLater(() -> {
    // Muestra la imagen en un JFrame
    JFrame marco = new JFrame();
    marco.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);
    marco.add(image);
//ajustar tamaño de marco
marco.setSize(1300,670);
marco.setLocationRelativeTo(null);
marco.setVisible(true);});
}    
 
/*public static void mostrarImagen(BufferedImage image, List<JFrame> frames) {
    // Muestra la imagen en un JFrame
    JFrame marco = new JFrame();
    marco.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);
    JLabel etiqueta = new JLabel(new ImageIcon(image));
    marco.getContentPane().add(etiqueta, BorderLayout.CENTER);
    marco.pack();
    marco.setVisible(true);
    // Añade el marco a la lista de marcos
    frames.add(marco);
}   */ 
     public static boolean buscar(String nombre,Sheet hoja, int columna) {
         Iterator<Row> iterador = hoja.iterator();
         while (iterador.hasNext()) {
             Row sigfila = iterador.next();
             Cell celnombre = sigfila.getCell(columna);
             if (celnombre != null && celnombre.getStringCellValue().equalsIgnoreCase(nombre)) {
                 System.out.println(azul+"Registrado existente. Aquí están los detalles:");
                 System.out.println("Marca: " + sigfila.getCell(0).getStringCellValue());
                 System.out.println("Modelo: " + sigfila.getCell(1).getStringCellValue());
                Cell cell = sigfila.getCell(2); // obtén la celda
                if (cell != null) {
                switch (cell.getCellType()) {
                    case STRING:
                    System.out.println("Año: " + sigfila.getCell(2).getStringCellValue());
                    
                        break;
                    case NUMERIC:
                    System.out.println("Año: " + sigfila.getCell(2).getNumericCellValue());
                       
                        break;
                    default:
                        System.out.println("Tipo de celda no soportado");
                }
            }else {
                System.out.println("La celda es null");
            }
                 System.out.println("Descripcion: " + sigfila.getCell(3).getStringCellValue()+reset);
                 return true; // devuelve true si encuentra un registro existente
 
             }
         }
         return false; // devuelve false si no encuentra un registro existente
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
     private static int campo(Scanner infu,String mensaje) { 
    
        int ano = 0;
        do {
            System.out.println(mensaje);
            while (!infu.hasNextInt()) {
                System.out.println(rojo+"Campo Obligatorio. Por favor, introduce un número."+reset);
                infu.next(); // descarta la entrada incorrecta
            }
            ano = infu.nextInt();
            infu.nextLine(); // consume newline
        } while (ano <= 0);
        return ano;
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
     private static String leerruta(Scanner infu, String mensaje) {
        String dato = ""; // ruta del archivo
        do {
            System.out.println(mensaje);
            try {
                dato = infu.nextLine().trim();
                if (!dato.matches("[a-zA-Z0-9_/\\\\\\\\.\\\\(\\\\)\\\\s-]+")) { // si la entrada contiene algo más que letras, números, barras, barras invertidas y puntos
                    throw new IllegalArgumentException(rojo + "Por favor, introduce una ruta de archivo válida." + reset);
                }
            } catch (IllegalArgumentException e) {
                System.out.println(rojo + e.getMessage() + reset);
                dato = ""; // resetea 'dato' para que el bucle continue
            }
            if (dato.isEmpty()) {
                System.out.println(rojo + "Campo Obligatorio." + reset);
            }
        } while (dato.isEmpty());
        return dato;
    }
    
     private static String ruta(Scanner infu, String mensaje) {
        String rutaImagen = "";
        boolean imagenCargada = false;
        while (!imagenCargada) {
            try {
                rutaImagen = leerruta(infu, mensaje);
                InputStream inputStream = new FileInputStream(rutaImagen);
                inputStream.close();
                imagenCargada = true;
            } catch (Exception e) {
                System.out.println("Error al cargar la imagen. Por favor, verifica la ruta de la imagen.");
            }
        }
        return rutaImagen;
    }
    
     public static void eliminar(String marca, String modelo, Sheet hoja, int columnamarc, int columnamod,List<JFrame> frames) {
        String Descripcion="";
        int Ano=0;
        List<Row> rowsToRemove = new ArrayList<>();
        Iterator<Row> iterador = hoja.iterator();
        while (iterador.hasNext()) {
            Row sigfila = iterador.next();
            int numerofila= sigfila.getRowNum();
            Cell celmarca = sigfila.getCell(columnamarc);
            Cell celmodelo = sigfila.getCell(columnamod);
            if (celmarca != null && celmarca.getStringCellValue().equalsIgnoreCase(marca) && celmodelo != null && celmodelo.getStringCellValue().equalsIgnoreCase(modelo)) {
                // Buscar la imagen
                BufferedImage image=null,image1=null,image2=null;
                for(int i=4;i<=6;i++){
                    if(i==4){
                        image= buscarImagen(hoja, sigfila,i);
                    }else if(i==5){
                        image1 = buscarImagen(hoja, sigfila,i);
                    }else if(i==6){
                        image2 = buscarImagen(hoja, sigfila,i);
                    }
                }
                if (image != null || image1 != null || image2 != null) {
                    // Crear una instancia de MiPanel
                    
                    tablaregistro1 panel = new tablaregistro1(image, image1,image2,marca, modelo,Ano, Descripcion);
                    
                    // Crear un JFrame para mostrar el panel
                    JFrame frame = new JFrame();
                    frame.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);
                    frame.add(panel);
                    frame.setSize(500,500);
                    frame.setLocationRelativeTo(null);
                    frame.setVisible(true);
    
                    // Añadir el marco a la lista de marcos
                    frames.add(frame);
                }
                if (image != null || image1 != null || image2 != null) {
                    // Mostrar un mensaje de confirmación
                    int confirm = JOptionPane.showConfirmDialog(null, "¿Deseas eliminar esta fila y sus imágenes asociadas?", "Confirmación", JOptionPane.YES_NO_OPTION);
                    if (confirm == JOptionPane.YES_OPTION) {
                        // Agregar la fila a la lista de filas a eliminar
                        rowsToRemove.add(sigfila);
                    }
                }
            }
        }
    
        // Eliminar las filas de la hoja
        for (Row row : rowsToRemove) {
            int filaeliminar = row.getRowNum();
            hoja.removeRow(row);
            System.out.println("La fila ha sido eliminada.");
            
            // Mover todas las filas debajo de la fila eliminada una posición hacia arriba
            int ultimafila = hoja.getLastRowNum();
            if (filaeliminar >= 0 && filaeliminar < ultimafila) {
                hoja.shiftRows(filaeliminar + 1, ultimafila, -1);
            }
    
    // Ajustar las imágenes de las filas que se han movido
    XSSFDrawing drawing = ((XSSFSheet) hoja).createDrawingPatriarch();
    for (XSSFShape shape : drawing.getShapes()) {
        if (shape instanceof XSSFPicture) {
            XSSFPicture picture = (XSSFPicture) shape;
            ClientAnchor anchor = picture.getPreferredSize();
            if (anchor.getRow1() > filaeliminar&& anchor.getRow1() < ultimafila&& anchor.getRow1() != filaeliminar && anchor.getRow1() != ultimafila) {
                // Mover la imagen una posición hacia arriba
                anchor.setRow1(anchor.getRow1() - 1);
                anchor.setRow2(anchor.getRow2() - 1);
            }
        }
    }
           
            //Añade lista para las imagenes
            List<XSSFPicture> imageneseliminar = new ArrayList<>();
            for (XSSFShape shape : drawing.getShapes()) {
                if (shape instanceof XSSFPicture) {
                    XSSFPicture picture = (XSSFPicture) shape;
                    ClientAnchor anchor = picture.getPreferredSize();
                    if (anchor.getRow1() == filaeliminar) {
                        imageneseliminar.add(picture);
                    }
                }
            }
            for (XSSFPicture picture : imageneseliminar) {
                deleteCTAnchor(picture);
                System.out.println("La imagen ha sido eliminada.");
            }
        }
    }
    public static void deleteCTAnchor(XSSFPicture picture) {
        XSSFDrawing drawing = picture.getDrawing();
        XmlCursor cursor = picture.getCTPicture().newCursor();
        cursor.toParent();
        if (cursor.getObject() instanceof org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor) {
            for (int i = 0; i < drawing.getCTDrawing().getTwoCellAnchorList().size(); i++) {
                if (cursor.getObject().equals(drawing.getCTDrawing().getTwoCellAnchorArray(i))) {
                    drawing.getCTDrawing().removeTwoCellAnchor(i);
                    System.out.println("Se eliminó el CTTwoCellAnchor de la imagen.");
                }
            }
        } else if (cursor.getObject() instanceof org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor) {
            for (int i = 0; i < drawing.getCTDrawing().getOneCellAnchorList().size(); i++) {
                if (cursor.getObject().equals(drawing.getCTDrawing().getOneCellAnchorArray(i))) {
                    drawing.getCTDrawing().removeOneCellAnchor(i);
                    System.out.println("Se eliminó el CTOneCellAnchor de la imagen.");
                }
            }
        }
    }
}