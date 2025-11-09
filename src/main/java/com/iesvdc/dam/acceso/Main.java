package com.iesvdc.dam.acceso;

import java.io.File;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Scanner;

import com.iesvdc.dam.acceso.conexion.Conexion;

public class Main {
    private static final Scanner sc = new Scanner(System.in);

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

    private static void mostrarCabecera() {
        System.out.println("╔════════════════════════════════════════════════════════╗");
        System.out.println("║                                                        ║");
        System.out.println("║        APLICACIÓN DE CONVERSIÓN EXCEL ↔ DATABASE       ║");
        System.out.println("║                                                        ║");
        System.out.println("╚════════════════════════════════════════════════════════╝\n");
    }

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

    private static void mostrarCabeceraOperacion(String titulo) {
        System.out.println("\n╔════════════════════════════════════════════════════════╗");
        System.out.printf ("║ %-54s ║%n", titulo);
        System.out.println("╚════════════════════════════════════════════════════════╝");
    }

    private static void mostrarMensajeExito(String msg) {
        System.out.println("\n╔════════════════════════════════════════════════════════╗");
        System.out.println("║                      OPERACIÓN EXITOSA                 ║");
        System.out.println("╠════════════════════════════════════════════════════════╣");
        System.out.printf ("║ %-54s ║%n", msg);
        System.out.println("╚════════════════════════════════════════════════════════╝");
    }

    private static void mostrarMensajeError(String msg) {
        System.out.println("\n╔════════════════════════════════════════════════════════╗");
        System.out.println("║                      ERROR DETECTADO                   ║");
        System.out.println("╠════════════════════════════════════════════════════════╣");
        System.out.printf ("║ %-54s ║%n", msg);
        System.out.println("╚════════════════════════════════════════════════════════╝");
    }

    private static void mostrarDespedida() {
        System.out.println("\n╔════════════════════════════════════════════════════════╗");
        System.out.println("║                                                        ║");
        System.out.println("║        Gracias por usar la aplicación de prueba         ║");
        System.out.println("║              Excel ↔ Base de Datos (IESVDC)             ║");
        System.out.println("║                                                        ║");
        System.out.println("╚════════════════════════════════════════════════════════╝");
    }

    private static int leerOpcion() {
        try {
            return Integer.parseInt(sc.nextLine().trim());
        } catch (NumberFormatException e) {
            return -1;
        }
    }

    private static void limpiarPantalla() {
        for (int i = 0; i < 30; i++) System.out.println();
    }
    //datos\PersonasSimp.xlsx
    
}
