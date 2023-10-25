package org.utils;

import org.utils.configuration.historicoCarteraConsumoPorOF.HistoricoCarteraConsumoPorOF;
import org.utils.configuration.historicoCarteraComercialPorOF.HistoricoCarteraComercialPorOF;
import org.utils.configuration.historicoCarteraMicrocreditoPorOF.HistoricoCarteraMicrocreditoPorOF;


import javax.swing.*;
import java.io.File;

import static org.utils.MethotsAzureMasterFiles.getDocument;
import static org.utils.MethotsAzureMasterFiles.*;

public class Start {


    public static void start() {
        System.out.println("\n" +
                "  _______   ___      _________________________.____     \n" +
                " /   ___/  /  _  \\    /     \\__    _/\\_   ___/|    |    \n" +
                " \\_____  \\  /  /_\\  \\  /  \\ /  \\|    |    |    _) |    |    \n" +
                " /        \\/    |    \\/    Y    \\    |    |        \\|    |___ \n" +
                "/_______  /\\____|__  /\\____|__  /____|   /_______  /|_______ \\\n" +
                "        \\/         \\/         \\/                 \\/         \\/\n");
        System.out.println("BIENVENIDO, VAMOS A REALIZAR UN TEST DE LA DATA");
        System.out.println("Espere por favor, va iniciar el proceso");
        try {
            //Ponemos a "Dormir" el programa 5sg
            Thread.sleep(5 * 1000);
            System.out.println("Generando analisis...");
            System.console();
            excecution();
            runtime();
        } catch (Exception e) {
            System.out.println(e);
        }
    }
    public static void excecution(){
        JOptionPane.showMessageDialog(null, "Seleccione el archivo Maestro");
        String masterFile = getDocument();

        try {
            assert masterFile != null;
            File file = new File(masterFile);
            System.out.println(file.getName());
            String fileName = file.getName().toLowerCase();
            System.out.println(fileName);
            if (fileName.contains("comercial")){
                HistoricoCarteraComercialPorOF.configuracion(masterFile);
            } else if (fileName.contains("consumo")) {
                HistoricoCarteraConsumoPorOF.configuracion(masterFile);
            } else if (fileName.contains("microcredito")) {
                HistoricoCarteraMicrocreditoPorOF.configuracion(masterFile);
            }else {
                System.out.println("EL ARCHIVO SELECCIONADO NO TIENE ANÁLISIS ASIGNADO");
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }
}
