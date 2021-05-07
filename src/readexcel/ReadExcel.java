/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package readexcel;

import java.io.File;
import java.io.IOException;

/**
 *
 * @author Andy
 */
public class ReadExcel {

    /**
     * @param args the command line arguments
     * @throws java.io.IOException
     */
    public static void main(String[] args) throws IOException {
        File excel = new File("C:\\Users\\Andy\\Desktop\\proyectos Telefonica\\documentacion general\\extracción.xlsx");

        Lectorexcel ejecutar = new Lectorexcel();
        ejecutar.LeerArchivo(excel);
       // System.out.println(ejecutar.getTransanccion().get(3));

    }

}
