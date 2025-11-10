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
/**
 * Clase de utilidades para trabajar con celdas de Excel usando Apache POI.
 * 
 * <p>Incluye métodos para determinar el tipo de dato de una celda y para 
 * establecer valores de celdas a partir de un {@link java.sql.ResultSet}.</p>
 * 
 * <p>Esta clase está pensada para facilitar la conversión entre datos 
 * provenientes de una base de datos y su representación en hojas de cálculo Excel.</p>
 * 
 * @author Antonio Martinez
 * @version 1.0
 */
public class ExcelUtils {
     /**
     * Margen de error utilizado para comparar números de tipo {@code double}.
     * Se usa para determinar si un valor numérico es entero o decimal.
     */
    private static final double EPSILON = 1e-10;

    /**
     * Determina el tipo SQL equivalente del contenido de una celda de Excel.
     * 
     * <p>Devuelve una cadena con el tipo de dato SQL más adecuado según el 
     * contenido de la celda. Si la celda es nula o está vacía, devuelve "Vacía".</p>
     * 
     * @param cell la celda de Excel a analizar (puede ser {@code null})
     * @return una cadena representando el tipo de dato SQL estimado:
     *         <ul>
     *             <li><b>varchar(255)</b> si el contenido es texto</li>
     *             <li><b>int</b> si el valor numérico es entero</li>
     *             <li><b>float</b> si el valor numérico es decimal</li>
     *             <li><b>date</b> si el valor numérico representa una fecha</li>
     *             <li><b>boolean</b> si el valor es booleano</li>
     *             <li><b>Fórmula</b> si la celda contiene una fórmula</li>
     *             <li><b>Vacía</b> si no contiene datos</li>
     *             <li><b>Error</b> si la celda presenta un error</li>
     *             <li><b>Desconocido</b> para otros casos no contemplados</li>
     *         </ul>
     */
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
    /**
     * Escribe el valor correspondiente de un {@link ResultSet} en una celda de una fila Excel.
     * 
     * <p>Según el tipo de campo especificado por {@link FieldType}, obtiene el valor del 
     * {@code ResultSet} y lo coloca en la celda indicada dentro de la fila.</p>
     * 
     * @param row         la fila de Excel donde se escribirá el valor
     * @param colIndex    el índice de columna (base 0) donde se creará la celda
     * @param rs          el {@link ResultSet} desde el cual se obtendrá el valor
     * @param columnName  el nombre de la columna en el {@code ResultSet}
     * @param fieldType   el tipo de dato del campo según {@link FieldType}
     * @throws SQLException si ocurre un error al acceder al {@code ResultSet}
     */
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
