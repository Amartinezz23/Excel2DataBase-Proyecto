package com.iesvdc.dam.acceso.modelo;

import java.util.List;


/**
 * El modelo que almacena el libro o lista de tablas.
 */
public class WorkbookModel {
    private final List<TableModel> tables;

    public List<TableModel> getTables() {
        return tables;
    }

    public WorkbookModel(List<TableModel> tables) {
        this.tables = tables;
    }
    
    
}
