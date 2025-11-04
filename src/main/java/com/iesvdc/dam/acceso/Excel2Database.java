package com.iesvdc.dam.acceso;

import java.sql.Connection;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.iesvdc.dam.acceso.conexion.Conexion;
import com.iesvdc.dam.acceso.excelutil.ExcelUtils;

/**
 * Este programa genérico en java (proyecto Maven) es un ejercicio 
 * simple que vuelca un libro Excel (xlsx) a una base de numeroCeldas (MySQL) 
 * y viceversa. El programa lee la configuración de la base de numeroCeldas 
 * de un fichero "properties" de Java y luego, con apache POI, leo 
 * las hojas, el nombre de cada hoja será el nombre de las tablas, 
 * la primera fila de cada hoja será el nombre de los atributos de 
 * cada tabla (hoja) y para saber el tipo de dato, tendré que 
 * preguntar a la segunda fila qué tipo de dato tiene. 
 * 
 * Procesamos el fichero Excel y creamos una estructura de numeroCeldas 
 * con la información siguiente: La estructura principal es el libro, 
 * que contiene una lista de tablas y cada tabla contiene tuplas 
 * nombre del campo y tipo de dato.
 *
 */
public class Excel2Database {
    
    public static void ExcelToDataBase(Connection conexion, String path){
        Map<Integer, String> mapSQL1 = new HashMap<>();
        try {
            conexion.setAutoCommit(false);

            Workbook workbook = new XSSFWorkbook(path);
            StringBuilder stringBuilderCreateTable = new StringBuilder();
            StringBuilder strimBuilderInsert = new StringBuilder();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                stringBuilderCreateTable =  getNombreTablaCreateTable(i, workbook);
                strimBuilderInsert = getNombreTablaInsert(i, workbook).append("( ");
                
                System.out.println();
                //System.out.println(strimBuilderInsert.toString());
                //System.out.println(stringBuilder.toString());
                Sheet sheet = workbook.getSheetAt(i);
                Row headers = sheet.getRow(0);
                Row tipoDato = sheet.getRow(1);
                int numeroCeldas = sheet.getLastRowNum();
                
                
                

                for (int j = 0; j < headers.getLastCellNum(); j++) {
                    stringBuilderCreateTable.append(headers.getCell(j).getStringCellValue()).append(" ").append(ExcelUtils.getTipoDato(tipoDato.getCell(j)));

                    strimBuilderInsert.append(headers.getCell(j).getStringCellValue());

                    // Si no es el último campo, añade coma
                    if (j != headers.getLastCellNum() - 1) {
                        stringBuilderCreateTable.append(",\n");
                        strimBuilderInsert.append(", ");
                    } else {
                        stringBuilderCreateTable.append("\n);");
                        strimBuilderInsert.append(")");
                    }
                    mapSQL1.put(i, stringBuilderCreateTable.toString());
                    //listaStrings.add(stringBuilderCreateTable.toString());
                }
                 Statement stmt = conexion.createStatement();
                 stmt.executeUpdate(mapSQL1.get(i));
                 strimBuilderInsert.append("\n").append("VALUES ");
                
                
               for (int r = 2; r <= sheet.getLastRowNum(); r++) { 
                    Row fila = sheet.getRow(r);
                    if (fila == null) continue; // evita filas vacías

                    strimBuilderInsert.append("(");

                    for (int c = 0; c < headers.getLastCellNum(); c++) {
                        Cell cell = fila.getCell(c);
                        String valor = "NULL";

                        if (cell != null) {
                            String tipo = ExcelUtils.getTipoDato(cell); // obtenemos el tipo usando tu método

                            switch (tipo) {
                                case "varchar(255)":
                                case "Fórmula": // si quieres tratar fórmulas como texto
                                    valor = "'" + cell.getStringCellValue() + "'";
                                    break;
                                case "int":
                                    valor = String.valueOf((int) cell.getNumericCellValue());
                                    break;
                                case "float":
                                    valor = String.valueOf(cell.getNumericCellValue());
                                    break;
                                case "date":
                                    valor = "'" + cell.getDateCellValue() + "'";
                                    break;
                                case "boolean":
                                    valor = String.valueOf(cell.getBooleanCellValue());
                                    break;
                                default:
                                valor = "NULL";
                            }
                        }

                        strimBuilderInsert.append(valor);

                        if (c != headers.getLastCellNum() - 1) {
                            strimBuilderInsert.append(", ");
                        } else {
                            strimBuilderInsert.append(")");
                        }
                }

                if (r != sheet.getLastRowNum()) {
                    strimBuilderInsert.append(",\n");
                } else {
                    strimBuilderInsert.append(";");
                }
                
                }
                
                
                
                String datos = strimBuilderInsert.toString();
                String arrayDatos[] = datos.split("(?<=;)");

                for (int b = 0; b < arrayDatos.length; b++) {
                    String insertSQL = arrayDatos[b].trim();
                    if (!insertSQL.isEmpty()) {
                        Statement stmt2 = conexion.createStatement();
                        stmt2.executeUpdate(insertSQL);  
                    }           
                }
                
            }
            conexion.commit();
            
            
            
        } catch (Exception e) {
            try {
                conexion.rollback();
            } catch (SQLException ex) {
                ex.printStackTrace();
            }
            e.printStackTrace();
        }finally {
            try {
                conexion.setAutoCommit(true);
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }


    public static StringBuilder getNombreTablaCreateTable(int num, Workbook workbook){
        StringBuilder builder = new StringBuilder();
        builder.append("CREATE TABLE ").append(workbook.getSheetName(num)).append(" (\n");

        return builder;
    }

    public static StringBuilder getNombreTablaInsert(int num, Workbook workbook){
        StringBuilder builder = new StringBuilder();
        builder.append("INSERT INTO ").append(workbook.getSheetName(num));

        return builder;
    }
    

    


}
/*
CREATE TABLE Persons (
    PersonID int,
    LastName varchar(255),
    FirstName varchar(255),
    Address varchar(255),
    City varchar(255)
);


INSERT INTO table_name (column1, column2, column3, ...)
VALUES (value1, value2, value3, ...);
 */

