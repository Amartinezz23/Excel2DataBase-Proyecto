package com.iesvdc.dam.acceso.excelutil;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

import com.iesvdc.dam.acceso.modelo.FieldType;

public class ExcelUtils {

    private static final double EPSILON = 1e-10;


    public static String getTipoDato(Cell cell) {
        if (cell == null) {
            return "Vacía";
        }

        switch (cell.getCellType()) {
            case STRING:
                return "varchar(255)";

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return "date";
                } else {
                    double valor = cell.getNumericCellValue();
                    if (Math.abs(valor - Math.floor(valor)) < EPSILON) {
                        return "int";
                    } else {
                        return "float";
                    }
                }

            case BOOLEAN:
                return "boolean";

            case FORMULA:
                // Puedes decidir si quieres evaluar la fórmula o solo indicar que es fórmula
                return "Fórmula";

            case BLANK:
                return "Vacía";

            case ERROR:
                return "Error";

            default:
                return "Desconocido";
        }
    }

    public static void setCellValueFromResultSet(Row row, int colIndex, ResultSet rs, 
                                               String columnName, FieldType fieldType) throws SQLException {
        switch (fieldType) {
            case INTEGER:
                int intValue = rs.getInt(columnName);
                if (!rs.wasNull()) {
                    row.createCell(colIndex).setCellValue(intValue);
                }
                break;
            case DECIMAL:
                double doubleValue = rs.getDouble(columnName);
                if (!rs.wasNull()) {
                    row.createCell(colIndex).setCellValue(doubleValue);
                }
                break;
            case DATE:
                java.sql.Date dateValue = rs.getDate(columnName);
                if (dateValue != null) {
                    row.createCell(colIndex).setCellValue(dateValue.toString());
                }
                break;
            case BOOLEAN:
                boolean boolValue = rs.getBoolean(columnName);
                if (!rs.wasNull()) {
                    row.createCell(colIndex).setCellValue(boolValue);
                }
                break;
            case STRING:
            default:
                String stringValue = rs.getString(columnName);
                if (stringValue != null) {
                    row.createCell(colIndex).setCellValue(stringValue);
                }
                break;
        }
    }
}
