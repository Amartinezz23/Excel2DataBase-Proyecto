package com.iesvdc.dam.acceso;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import com.iesvdc.dam.acceso.conexion.Conexion;

public class Main {
    public static void main(String[] args) {
        Connection connection = Conexion.getConnection();
        //Excel2Database.ExcelToDataBase(connection, "datos\\personas.xlsx");
        DataBase2Excel_2.generarExcel("datos\\personasOut.xlsx");
    }
}
