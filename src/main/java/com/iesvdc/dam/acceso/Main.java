package com.iesvdc.dam.acceso;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import com.iesvdc.dam.acceso.conexion.Conexion;

public class Main {
    public static void main(String[] args) {
        Connection connection = Conexion.getConnection();
        String sqlDataType = "SELECT DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE table_name = " + "'" + "ventas" + "';";
        try {
                
                
                
                Statement statementDataType = connection.createStatement();
                ResultSet resultSetDataType = statementDataType.executeQuery(sqlDataType);
                while (resultSetDataType.next()) {
                    
                    
                    String dataType = resultSetDataType.getString(1);
                    //datos.put(valorHeader, mapSQLTypeToFieldType(dataType));
                    System.out.println( dataType);
                    
                    //System.out.println(valorHeader);
                }
                //System.out.println("---------------------");
                 
            } catch (SQLException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
    }
}
