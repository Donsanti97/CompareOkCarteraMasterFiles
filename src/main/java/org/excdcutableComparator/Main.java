package org.excdcutableComparator;


import org.utils.Start;

import javax.swing.*;
import java.io.FileOutputStream;
import java.io.PrintStream;

import static org.utils.MethotsAzureMasterFiles.*;

public class Main {
    public static void main(String[] args) {

        Start.start();

        // Guarda la salida de error actual
        /*PrintStream originalErr = System.err;

        try {
            JOptionPane.showMessageDialog(null, "Seleccione a continuación la ubicación donde desea " +
                    "\n quede el documento para errores");
            String documentoErrores = getDirectory() + "\\Errores.txt";
            // Redirige la salida de error a un archivo
            FileOutputStream fileOutputStream = new FileOutputStream(documentoErrores);
            PrintStream printStream = new PrintStream(fileOutputStream);
            System.setErr(printStream);

            // Tu código va aquí
            // ...

            // Simulación de un error para demostrar la redirección
            throw new RuntimeException("Este es un mensaje de error de ejemplo.");

        } catch (Exception e) {
            // Manejar la excepción si es necesario
            e.printStackTrace();
        } finally {
            // Restaura la salida de error original
            System.setErr(originalErr);
        }*/
    }
}