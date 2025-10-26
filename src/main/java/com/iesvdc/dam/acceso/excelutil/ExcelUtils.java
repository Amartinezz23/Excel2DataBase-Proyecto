package com.iesvdc.dam.acceso.excelutil;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

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
}
