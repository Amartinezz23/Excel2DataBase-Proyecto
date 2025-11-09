package com.iesvdc.dam.acceso;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.iesvdc.dam.acceso.conexion.Conexion;
import com.iesvdc.dam.acceso.excelutil.ExcelUtils;
import com.iesvdc.dam.acceso.modelo.FieldType;
import com.iesvdc.dam.acceso.modelo.TableModel;
import com.iesvdc.dam.acceso.modelo.WorkbookModel;

public class DataBase2Excel_2 {

    

    

    public static FieldType mapSQLTypeToFieldType(String sqlType) {
        if (sqlType == null)
            return FieldType.UNKNOWN;

        sqlType = sqlType.toUpperCase();

        if (sqlType.contains("INT") || sqlType.equals("BIGINT") || sqlType.equals("SMALLINT")) {
            return FieldType.INTEGER;
        } else if (sqlType.contains("DECIMAL") || sqlType.contains("NUMERIC") || sqlType.contains("FLOAT")
                || sqlType.contains("DOUBLE")) {
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

    

    public static void generarExcel(String outPath){
        XSSFWorkbook workbook = new XSSFWorkbook();
        Connection connection = Conexion.getConnection();

        HashMap<String, LinkedHashMap<String,FieldType>> mapa = obtenerEstructura(connection);
        for (String tabla : mapa.keySet()) {
            Sheet hoja = workbook.createSheet(tabla);
             LinkedHashMap<String, FieldType> columas = mapa.get(tabla);
             int contador = 0;
             Row cabecera = hoja.createRow(0);
             for(String columna : columas.keySet()){
                cabecera.createCell(contador).setCellValue(columna);
                contador++;
             }
             introducirDatosExcel(connection, hoja, tabla, columas);

        }
        OutputStream file;
        try {
            file = new FileOutputStream(outPath);
            try {
                workbook.write(file);
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        
    }

    public static HashMap<String,LinkedHashMap<String,FieldType>> obtenerEstructura(Connection connection){
        HashMap<String, LinkedHashMap<String, FieldType>> mapaEstrutura = new HashMap<>();
        try {
            DatabaseMetaData data = connection.getMetaData();
            String catalogo = connection.getCatalog();
            //Pasamos el catalogo, null en el schema porque nos da igual, que nos muestre cualquiera, % para que nos de TODAS las tablas y luego 
            //"TABLE" que es un array de tablas, el tipo de objeto que queremos de la base de datos
            ResultSet resultSet = data.getTables(catalogo, null, "%", new String[]{"TABLE"});
            while (resultSet.next()) {
                String nombreTabla = resultSet.getString("TABLE_NAME");
                LinkedHashMap<String, FieldType> columnas = new LinkedHashMap<>();
                //Pasamos el catalogo, no queremos eschema de nuevo, nombre de la tabla y % para conocer todas las columnas de la tabla
                ResultSet set = data.getColumns(catalogo, null, nombreTabla, "%");
                while (set.next()) {
                    String nombreColumna = set.getString("COLUMN_NAME");
                    String tipoColumna = set.getString("TYPE_NAME");
                    FieldType tipoExcel = mapSQLTypeToFieldType(tipoColumna);
                    columnas.put(nombreColumna, tipoExcel);                    
                }
                set.close();
                mapaEstrutura.put(nombreTabla, columnas);
            }
            resultSet.close();
        } catch (Exception e) {
            // TODO: handle exception
        }
        
        return mapaEstrutura;
    }

    public static void introducirDatosExcel(Connection connection, Sheet sheet , String tabla, LinkedHashMap<String,FieldType> hashMap){
        StringBuilder builder = new StringBuilder();
        String[] nombresColumnas = hashMap.keySet().toArray(new String[0]);
        builder.append("SELECT ");
        
        for (String linea : hashMap.keySet()) {
            builder.append(linea).append(",");
            
        }
        builder.deleteCharAt(builder.length() -1);
        builder.append(" FROM ").append(tabla).append(";");

        String sql = builder.toString();

        int filaNumero = 1;
        try {
            Statement statement = connection.createStatement();
            ResultSet resultSet = statement.executeQuery(sql);
            while (resultSet.next()) {
                Row row = sheet.createRow(filaNumero);
                for (int i = 0; i < nombresColumnas.length; i++) {
                    String nombreCampo = nombresColumnas[i];
                    FieldType tipoCampo = hashMap.get(nombreCampo);
                    ExcelUtils.setCellValueFromResultSet(row, i, resultSet, nombreCampo, tipoCampo);
                }
                filaNumero++;
            }
        } catch (Exception e) {
            // TODO: handle exception
        }
        
        
        


    }

    

    

    

    
    public static void main(String[] args) {
        Connection connection = Conexion.getConnection();
        generarExcel("datos\\outFileExcel.xlsx");
        
        
    }

}
