package org.utils;

import org.utils.configuration.historicoCarteraBrutaPorOF_LC.HistoricoCarteraBrutaPorOF_LC;
import org.utils.configuration.historicoCarteraComercialPorOF.HistoricoCarteraComercialPorOF;
import org.utils.configuration.historicoCarteraConsumoPorOF.HistoricoCarteraConsumoPorOF;
import org.utils.configuration.historicoCarteraMicrocreditoPorOF.HistoricoCarteraMicrocreditoPorOF;
import org.utils.configuration.historicoCarteraPorLC.HistoricoCarteraPorLC;
import org.utils.configuration.historicoCarteraPorOF.HistoricoCarteraPorOF;
import org.utils.configuration.historicoCarteraSegMonto_ColocPorLC.HistoricoCarteraSegMonto_ColocPorLC;
import org.utils.configuration.historicoCarteraSegMonto_ColocPorOF.HistoricoCarteraSegMonto_ColocPorOF;

import javax.swing.*;
import java.io.File;

import static org.utils.MethotsAzureMasterFiles.getDocument;
import static org.utils.MethotsAzureMasterFiles.runtime;

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

    public static void excecution() {
        JOptionPane.showMessageDialog(null, "Seleccione el archivo Maestro");
        String masterFile = getDocument();
        while (masterFile ==null){
            JOptionPane.showMessageDialog(null, "NO HA SELECCIONADO NINGÚN ARCHIVO " +
                    "\n POR FAVOR SELECCIONE UN ARCHIVO MAESTRO A ANALIZAR");
            masterFile = getDocument();
        }

        try {
            File file = new File(masterFile);
            System.out.println(file.getName());
            String fileName = file.getName().toLowerCase();
            System.out.println(fileName);
            if (fileName.contains("comercial")) {
                HistoricoCarteraComercialPorOF.configuracion(masterFile);
            } else if (fileName.contains("consumo")) {
                HistoricoCarteraConsumoPorOF.configuracion(masterFile);
            } else if (fileName.contains("microcredito")) {
                HistoricoCarteraMicrocreditoPorOF.configuracion(masterFile);
            } else if (fileName.contains("historico cartera por lc")) {
                HistoricoCarteraPorLC.configuracion(masterFile);
            } else if (fileName.contains("seg")) {
                if (fileName.contains("lc")) {
                    HistoricoCarteraSegMonto_ColocPorLC.configuracion(masterFile);
                } else if (fileName.contains("of")) {
                    HistoricoCarteraSegMonto_ColocPorOF.configuracion(masterFile);
                }

            } else if (fileName.contains("historico cartera por of")) {
                HistoricoCarteraPorOF.configuracion(masterFile);
            } else if (fileName.contains("historico cartera bruta por of _ lc")) {
                HistoricoCarteraBrutaPorOF_LC.configuracion(masterFile);
            } else {
                System.out.println("EL ARCHIVO SELECCIONADO NO TIENE ANÁLISIS ASIGNADO");
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }
}
