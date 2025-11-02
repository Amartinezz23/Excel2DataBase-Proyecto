package com.iesvdc.dam.acceso;

import java.io.File;
import java.nio.file.Files;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.iesvdc.dam.acceso.conexion.Conexion;
import com.iesvdc.dam.acceso.modelo.FieldType;
import com.iesvdc.dam.acceso.modelo.TableModel;
import com.iesvdc.dam.acceso.modelo.WorkbookModel;

public class DataBase2Excel {

    

    public static List<String> obtenerTablas(Connection connection) {
        List<String> tablas = new ArrayList<>();
        ResultSet rs = null;
        try {
            DatabaseMetaData data = connection.getMetaData();
            rs = data.getTables(null, null, "%", new String[] { "TABLE" });
            while (rs.next()) {
                tablas.add(rs.getString("TABLE_NAME"));
            }
        } catch (Exception e) {
            e.getMessage();
        }
        try {
            rs.close();
        } catch (SQLException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        return tablas;
    }

    

    public static void imprimirMaps(LinkedHashMap<String, LinkedHashMap<String,FieldType>> estructura){
        for (Entry<String, LinkedHashMap<String, FieldType>> tablaEntry : estructura.entrySet()) {
                String tablaX = tablaEntry.getKey();
                LinkedHashMap<String, FieldType> columnas = tablaEntry.getValue();

                System.out.println("Tabla: " + tablaX);
                System.out.println("-------------------");

                for (Entry<String, FieldType> columnaEntry : columnas.entrySet()) {
                    String columna = columnaEntry.getKey();
                    FieldType tipo = columnaEntry.getValue();
                    System.out.printf("  %-15s : %s%n", columna, tipo); // Formato alineado
                }

                System.out.println(); // Línea en blanco entre tablas
            }
    }
    public static FieldType mapSQLTypeToFieldType(String sqlType) {
    if (sqlType == null) return FieldType.UNKNOWN;

    sqlType = sqlType.toUpperCase();

    if (sqlType.contains("INT") || sqlType.equals("BIGINT") || sqlType.equals("SMALLINT")) {
        return FieldType.INTEGER;
    } else if (sqlType.contains("DECIMAL") || sqlType.contains("NUMERIC") || sqlType.contains("FLOAT") || sqlType.contains("DOUBLE")) {
        return FieldType.DECIMAL;
    } else if (sqlType.contains("CHAR") || sqlType.contains("TEXT") || sqlType.contains("VARCHAR")) {
        return FieldType.STRING;
    } else if (sqlType.contains("DATE") || sqlType.contains("TIME")) {
        return FieldType.DATE;
    } else if (sqlType.equals("BIT") || sqlType.equals("BOOLEAN")) {
        return FieldType.BOOLEAN;
    } else {
        return FieldType.UNKNOWN;
    }
}
        private static LinkedHashMap<String, String> obtenerColumnasYTipos(Connection connection, String tabla) throws SQLException {
        LinkedHashMap<String, String> columnas = new LinkedHashMap<>();

        String sql = """
            SELECT COLUMN_NAME, DATA_TYPE 
            FROM INFORMATION_SCHEMA.COLUMNS 
            WHERE TABLE_SCHEMA = 'agenda' AND TABLE_NAME = '%s'
            """.formatted(tabla);

        try (Statement stmt = connection.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {

            while (rs.next()) {
                String nombreColumna = rs.getString("COLUMN_NAME");
                String tipoSQL = rs.getString("DATA_TYPE");
                columnas.put(nombreColumna, tipoSQL);
            }
        }

        return columnas;
    }

    public static LinkedHashMap<String, LinkedHashMap<String, FieldType>> obtenerEstructuraTablas(
            Connection connection, List<String> tablas) {

        LinkedHashMap<String, LinkedHashMap<String, FieldType>> estructura = new LinkedHashMap<>();

        for (String tabla : tablas) {
            LinkedHashMap<String, FieldType> columnas = new LinkedHashMap<>();
            try {
                LinkedHashMap<String, String> headersYTipos = obtenerColumnasYTipos(connection, tabla);
                for (var entry : headersYTipos.entrySet()) {
                    columnas.put(entry.getKey(), mapSQLTypeToFieldType(entry.getValue()));
                }
            } catch (SQLException e) {
                System.err.println("Error al obtener estructura de tabla '" + tabla + "': " + e.getMessage());
            }
            estructura.put(tabla, columnas);
        }

        return estructura;
    }

    public static void main(String[] args) {
        List<String> tablas = null;
        LinkedHashMap<String, LinkedHashMap<String, FieldType>> estructura = null;
        try (Connection connection = Conexion.getConnection()) {
            tablas = obtenerTablas(connection);
            estructura = obtenerEstructuraTablas(connection, tablas);
            imprimirMaps(estructura);
        } catch (SQLException e) {
            System.err.println("Error al conectar o procesar la base de datos: " + e.getMessage());
            e.printStackTrace();
        }

        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        int conator = 0;
        for ( String tabla : estructura.keySet()) {
            hssfWorkbook.createSheet(tabla);
            for (int i = 0; i < hssfWorkbook.getNumberOfSheets(); i++) {
                Sheet hoja = hssfWorkbook.getSheetAt(i);
                Row cabecera = hoja.getRow(0);
                for (LinkedHashMap mapValues : estructura.values()) {
                    mapValues.keySet().forEach(System.out::println);
                    conator++;
                    
                }
                System.out.println("--------------------------");
            }
            System.out.println("--------------------------");
        }
        
        
    }

    

}
