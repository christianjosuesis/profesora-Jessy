package com;
import java.util.Random;
import java.util.Scanner;
public class aleatorio {
    private int numero;//campo
    
    public aleatorio(int numero){//constructor
        this.numero=numero;
    }
       public boolean comparar(int numeroing){//metodo
        if(numeroing < this.numero){
            System.out.println("El numero es mayor.");
            return false;
        } else if(numeroing > this.numero) {
            System.out.println("El numero es menor");
            return false;
        } else {
            System.out.println("¡Felicidades! Has adivinado el número.");
            return true;
        }
    }
    public static void main (String args[]){
Scanner num=new Scanner(System.in);
String res="";
System.out.println("Este programa genera un numero aleatorio entre 1 y 50 el cual tienes que adivinar");
System.out.println();
System.out.println("Se a generado un numero intenta adivinar cual es te giare diciendo si es mayor o menor.");
System.out.println();  
do{
System.out.println("COMENZAMOS.");
Random numal = new Random(); 
int numeroAleatorio = numal.nextInt(50) + 1;
aleatorio aleatori= new aleatorio(numeroAleatorio);
int nume;
boolean aceptado;
do {
    System.out.print("El numero es: "); 
    while (!num.hasNextInt()) {//verifica si el próximo token en la entrada del usuario no puede interpretarse como un int. Si el próximo token no puede convertirse en un int, !num.hasNextInt() devuelve true. Si el próximo token puede convertirse en un int, !num.hasNextInt() devuelve false.
        System.out.println("El numero es un entero.");
    num.next(); // descarta la entrada incorrecta
}
nume=num.nextInt();
num.nextLine(); // consume newline
aceptado=aleatori.comparar(nume);
} while (!aceptado);
 while(true){
    System.out.println("Quieres volver a jugar. SI/NO ");
    res=num.nextLine().toUpperCase();
    if (!res.matches(".*\\d.*")) {
    }else{ 
        System.out.println("Por favor solo ingresa texto.");
         }
    if (res.equals("SI") || res.equals("NO")) {
    //.*: Este es un cuantificador en expresiones regulares que significa “cero o más de cualquier carácter”.
    // \\d: Este es un metacarácter en expresiones regulares que significa “cualquier dígito”.
}else{ 
    System.out.println("Entrada inválida. Por favor, ingresa SI o NO.");
     }
     if(res.equals("SI")){
        System.out.println("Se a generado un nuevo numero adivina cual es.");
        break;
     }
 }
}while(res.equalsIgnoreCase("SI")); 
}
}
