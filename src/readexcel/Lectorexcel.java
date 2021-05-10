/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package readexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Andy
 */
public class Lectorexcel {

    public List<String> Universo = new ArrayList();
    public List<String> Transanccion = new ArrayList();
    public List<String> Tabla = new ArrayList();
    public List<String> Fase = new ArrayList();
    public List<String> FechaIn = new ArrayList();
    public List<String> FechaFin = new ArrayList();
    public List<String> t_Systems =  new ArrayList();
    
    /*
    public List<String> Cnt_entradas = new ArrayList();
    public List<String> Conector_destino = new ArrayList();
    public List<String> Elemento_aprovis = new ArrayList();*/
    public int n_systems ;
    public String contenidoCelda, contenido_celda_sistemas;
    

    public Row row;
    /*Le damos la posicion inicial desde donde empiece a recorrer cada fila entre 0-10 en el excel es de 1-11
    y  en este caso desde las posicion 2
     */
    public int fila = 2;
    public int name_of_sys_row = 2;
    private final int ctn_system_fila =8;
    private final int ctn_system_colum = 1;
    
    

    public void LeerArchivo(File filename) throws FileNotFoundException, IOException {
        //Cargamos el archivo con la ruta  en la cual esta 
        FileInputStream ExcelFile = new FileInputStream(filename);
        
        // Generamos una referencia al Excel que vamos a leer
        XSSFWorkbook WB = new XSSFWorkbook(ExcelFile);
        // Le decimos la hoja que vamos a usar
        XSSFSheet sheet = WB.getSheetAt(0);
        //recorremos las 9filas  que vamos  usar 
        for (int nfila = 0; nfila < 6; nfila++) {
             System.out.println("\n");
            //indicicamos las columas que tiene que leer 
            for (int Colum = 1; Colum < 4; Colum++) {
                row = sheet.getRow(fila);
                //recoger los datos de la fila con sus respectivas columnas 
                contenidoCelda = row.getCell(Colum).getStringCellValue();
         //   System.out.println("celda: " + contenidoCelda);

                switch (fila) {
                    case 2:
                        Universo.add(contenidoCelda);
                        break;
                    case 3:
                        Transanccion.add(contenidoCelda);
                        break;
                    case 4:
                        Tabla.add(contenidoCelda);
                        break;
                    case 5:
                        Fase.add(contenidoCelda);
                        break;
                    case 6:
                        FechaIn.add(contenidoCelda);
                        break;
                    case 7:
                        FechaFin.add(contenidoCelda);
                        break;/*
                    case 8:
                        Cnt_entradas.add(contenidoCelda);
                        break;
                    case 9:
                        Conector_destino.add(contenidoCelda);
                        break;
                    case 10:
                        Elemento_aprovis.add(contenidoCelda);
                        break;*/
                    default:
                        System.out.println("No entra");

                }

            }

            //Aumenta cada fila  cuando acaba de leer cada una de las columnas
            fila++;

        }
        

    }
    
    public void Leer_systems (File filename) throws FileNotFoundException, IOException {
        
     //Cargamos el archivo con la ruta  en la cual esta 
        FileInputStream ExcelFile = new FileInputStream(filename);
        
        // Generamos una referencia al Excel que vamos a leer
        XSSFWorkbook WB = new XSSFWorkbook(ExcelFile);
        // Le decimos la hoja que vamos a usar
        XSSFSheet sheet = WB.getSheetAt(0);
        //recorremos las 9filas  que vamos  usar 
        
         row = sheet.getRow(ctn_system_fila);
                //recoger los datos de la fila con sus respectivas columnas 
        n_systems  = (int) row.getCell(ctn_system_colum).getNumericCellValue();
         
      //  System.out.println("celda: " + n_systems);
        
        for (int i = 0; i < n_systems; i++) {
            
            row = sheet.getRow(name_of_sys_row);
            
            contenido_celda_sistemas  = row.getCell(6).getStringCellValue();
            t_Systems.add(contenido_celda_sistemas);
           name_of_sys_row++; 
           //System.out.println("celda: " + contenido_celda_sistemas);
        }
}

    public List<String> getUniverso() {
        return Universo;
    }

    public void setUniverso(List<String> Universo) {
        this.Universo = Universo;
    }

    public List<String> getTransanccion() {
        return Transanccion;
    }

    public void setTransanccion(List<String> Transanccion) {
        this.Transanccion = Transanccion;
    }

    public List<String> getTabla() {
        return Tabla;
    }

    public void setTabla(List<String> Tabla) {
        this.Tabla = Tabla;
    }

    public List<String> getFase() {
        return Fase;
    }

    public void setFase(List<String> Fase) {
        this.Fase = Fase;
    }

    public List<String> getFechaIn() {
        return FechaIn;
    }

    public void setFechaIn(List<String> FechaIn) {
        this.FechaIn = FechaIn;
    }

    public List<String> getFechaFin() {
        return FechaFin;
    }

    public void setFechaFin(List<String> FechaFin) {
        this.FechaFin = FechaFin;
    }
/*
    public List<String> getCnt_entradas() {
        return Cnt_entradas;
    }

    public void setCnt_entradas(List<String> Cnt_entradas) {
        this.Cnt_entradas = Cnt_entradas;
    }

    public List<String> getConector_destino() {
        return Conector_destino;
    }

    public void setConector_destino(List<String> Conector_destino) {
        this.Conector_destino = Conector_destino;
    }

    public List<String> getElemento_aprovis() {
        return Elemento_aprovis;
    }

    public void setElemento_aprovis(List<String> Elemento_aprovis) {
        this.Elemento_aprovis = Elemento_aprovis;
    }
*/
    public String getContenidoCelda() {
        return contenidoCelda;
    }

    public void setContenidoCelda(String contenidoCelda) {
        this.contenidoCelda = contenidoCelda;
    }

    public Row getRow() {
        return row;
    }

    public void setRow(Row row) {
        this.row = row;
    }

    public int getFila() {
        return fila;
    }

    public void setFila(int fila) {
        this.fila = fila;
    }

    public List<String> getSystems() {
        return t_Systems;
    }

    public void setSystems(List<String> Systems) {
        this.t_Systems = Systems;
    }

    public int getN_systems() {
        return n_systems;
    }

    public void setN_systems(int n_systems) {
        this.n_systems = n_systems;
    }
    
    
}
