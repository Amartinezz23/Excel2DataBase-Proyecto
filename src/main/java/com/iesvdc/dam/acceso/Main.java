package com.iesvdc.dam.acceso;

import java.io.File;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Scanner;

import com.iesvdc.dam.acceso.conexion.Conexion;
/**
 * Clase principal de la aplicación de conversión entre Excel y Base de Datos.
 * 
 * <p>Permite realizar dos operaciones principales desde un menú interactivo por consola:</p>
 * <ul>
 *   <li><b>Importar</b> datos desde un archivo Excel a una base de datos.</li>
 *   <li><b>Exportar</b> datos desde la base de datos a un archivo Excel.</li>
 * </ul>
 * 
 * <p>También incluye opciones de salida y manejo básico de errores en consola.</p>
 * 
 * <p>Esta clase sirve como punto de entrada de la aplicación (método {@code main}).</p>
 * 
 * @author Antonio Martinez
 * @version 1.0
 */
public class Main {
    /**
     * Escáner utilizado para leer las entradas del usuario desde la consola.
     */
    private static final Scanner sc = new Scanner(System.in);

    /**
     * Método principal que inicia la ejecución de la aplicación.
     * 
     * <p>Muestra un menú interactivo con opciones para convertir datos 
     * entre Excel y la base de datos. El programa continúa ejecutándose 
     * hasta que el usuario elige la opción de salida.</p>
     * 
     * @param args argumentos de línea de comandos (no utilizados)
     */
    public static void main(String[] args) {
        boolean running = true;
        mostrarCabecera();

        while (running) {
            mostrarMenuPrincipal();
            int opcion = leerOpcion();

            switch (opcion) {
                case 1:
                    mostrarCabeceraOperacion("CONVERSIÓN DE EXCEL A BASE DE DATOS");
                    System.out.print("Ruta del archivo Excel (.xlsx): ");
                    String rutaExcel = sc.nextLine().trim();
                    if (rutaExcel.isEmpty()) rutaExcel = "datos/entrada.xlsx";
                    

                    File archivoExcel = new File(rutaExcel);
                    if (!archivoExcel.exists()) {
                        mostrarMensajeError("El archivo especificado no existe: " + rutaExcel);
                        break; 
                    }
                    System.out.println("\nProcesando archivo...");
                    try {
                        Excel2Database.ExcelToDataBase(Conexion.getConnection(), rutaExcel);
                        mostrarMensajeExito("Archivo Excel importado correctamente a la base de datos.");
                    } catch (Exception e) {
                        mostrarMensajeError("Error durante la conversión: " + e.getMessage());
                    }
                    break;

                case 2:
                    mostrarCabeceraOperacion("EXPORTAR BASE DE DATOS A EXCEL");
                    System.out.print("Ruta de salida (.xlsx): ");
                    String rutaSalida = sc.nextLine().trim();
                    if (rutaSalida.isEmpty()) rutaSalida = "datos/exportado.xlsx";
                    System.out.println("\nGenerando archivo Excel...");

                    try {
                        DataBase2Excel_2.generarExcel(rutaSalida);
                        mostrarMensajeExito("Base de datos exportada correctamente a " + rutaSalida);
                    } catch (Exception e) {
                        mostrarMensajeError("Error durante la exportación: " + e.getMessage());
                    }
                    break;

                case 3:
                    mostrarDespedida();
                    running = false;
                    break;

                default:
                    mostrarMensajeError("Opción inválida. Elija 1, 2 o 3.");
                    break;
            }

            if (running) {
                System.out.println("\nPresione Enter para continuar...");
                sc.nextLine();
                limpiarPantalla();
            }
        }

        sc.close();
    }

    /**
     * Muestra en consola la cabecera decorativa de la aplicación.
     */
    private static void mostrarCabecera() {
        System.out.println("╔════════════════════════════════════════════════════════╗");
        System.out.println("║                                                        ║");
        System.out.println("║        APLICACIÓN DE CONVERSIÓN EXCEL ↔ DATABASE       ║");
        System.out.println("║                                                        ║");
        System.out.println("╚════════════════════════════════════════════════════════╝\n");
    }
     /**
     * Muestra el menú principal con las opciones disponibles.
     */
    private static void mostrarMenuPrincipal() {
        System.out.println("╔════════════════════════════════════════════════════════╗");
        System.out.println("║                      MENÚ PRINCIPAL                    ║");
        System.out.println("╠════════════════════════════════════════════════════════╣");
        System.out.println("║  1. Importar desde Excel a Base de Datos               ║");
        System.out.println("║  2. Exportar desde Base de Datos a Excel               ║");
        System.out.println("║  3. Salir                                              ║");
        System.out.println("╚════════════════════════════════════════════════════════╝");
        System.out.print("\nSeleccione una opción [1-3]: ");
    }
    /**
     * Muestra una cabecera de operación personalizada con un título.
     * 
     * @param titulo título descriptivo de la operación actual
     */
    private static void mostrarCabeceraOperacion(String titulo) {
        System.out.println("\n╔════════════════════════════════════════════════════════╗");
        System.out.printf ("║ %-54s ║%n", titulo);
        System.out.println("╚════════════════════════════════════════════════════════╝");
    }
    /**
     * Muestra un mensaje de éxito en formato decorativo por consola.
     * 
     * @param msg el mensaje de éxito a mostrar
     */
    private static void mostrarMensajeExito(String msg) {
        System.out.println("\n╔════════════════════════════════════════════════════════╗");
        System.out.println("║                      OPERACIÓN EXITOSA                 ║");
        System.out.println("╠════════════════════════════════════════════════════════╣");
        System.out.printf ("║ %-54s ║%n", msg);
        System.out.println("╚════════════════════════════════════════════════════════╝");
    }
    /**
     * Muestra un mensaje de error en formato decorativo por consola.
     * 
     * @param msg el mensaje de error a mostrar
     */
    private static void mostrarMensajeError(String msg) {
        System.out.println("\n╔════════════════════════════════════════════════════════╗");
        System.out.println("║                      ERROR DETECTADO                   ║");
        System.out.println("╠════════════════════════════════════════════════════════╣");
        System.out.printf ("║ %-54s ║%n", msg);
        System.out.println("╚════════════════════════════════════════════════════════╝");
    }
    /**
     * Muestra un mensaje de despedida antes de finalizar la ejecución del programa.
     */
    private static void mostrarDespedida() {
        System.out.println("\n╔════════════════════════════════════════════════════════╗");
        System.out.println("║                                                        ║");
        System.out.println("║        Gracias por usar la aplicación de prueba         ║");
        System.out.println("║              Excel ↔ Base de Datos (IESVDC)             ║");
        System.out.println("║                                                        ║");
        System.out.println("╚════════════════════════════════════════════════════════╝");
    }
    /**
     * Lee una opción ingresada por el usuario desde la consola.
     * 
     * <p>Devuelve -1 si el valor ingresado no es numérico o no válido.</p>
     * 
     * @return el número de opción introducido, o -1 si ocurre un error
     */
    private static int leerOpcion() {
        try {
            return Integer.parseInt(sc.nextLine().trim());
        } catch (NumberFormatException e) {
            return -1;
        }
    }
    /**
     * Limpia parcialmente la consola imprimiendo varias líneas vacías.
     * 
     * <p>Esto no borra la consola realmente, pero mejora la legibilidad entre operaciones.</p>
     */
    private static void limpiarPantalla() {
        for (int i = 0; i < 30; i++) System.out.println();
    }
    //datos\PersonasSimp.xlsx
    
}
