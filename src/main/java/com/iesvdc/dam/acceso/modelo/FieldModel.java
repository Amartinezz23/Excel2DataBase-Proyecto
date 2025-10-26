package com.iesvdc.dam.acceso.modelo;

/**
 * El modelo que almacena información de un campo y sus propiedades.
 */

public class FieldModel {
    
    private final String name;
    private final FieldType type;
    
    public FieldModel(String name, FieldType type) {
        this.name = name;
        this.type = type;
    }

}
