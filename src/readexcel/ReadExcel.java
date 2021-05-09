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
        File excel = new File("C:\\Users\\Andy\\Documents\\NetBeansProjects\\Ex_Cmbs_Us_Mdt.xlsx");

        Lectorexcel ejecutar = new Lectorexcel();
        //ejecutar.LeerArchivo(excel);
        ejecutar.Leer_systems(excel);
       // System.out.println(ejecutar.getTransanccion().get(3));

    }

}
