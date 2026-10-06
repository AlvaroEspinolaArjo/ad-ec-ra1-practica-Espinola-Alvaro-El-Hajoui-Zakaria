package org.educa.service;

import jakarta.xml.bind.JAXBException;
import org.educa.dao.ProductoDAO;
import org.educa.entity.ProductoEntity;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.text.ParseException;
import java.util.List;

public class ProductoService {

    private ProductoDAO productoDAO = new ProductoDAO();

    public List<ProductoEntity> readFile(String fileXml) throws JAXBException {
        return productoDAO.readFile(fileXml); //llama al metodo readFile() de productoDAO y devuelve el resultado.
    }

    public void exportSummary(String path, String fileXml) throws JAXBException, IOException {

        // Leemos el fichero XML y obtenemos una lista con todos los productos
        List<ProductoEntity> productos = readFile(fileXml);

        // Creamos un objeto File que representa el fichero XML
        File ficheroXML = new File(fileXml);

        // Obtenemos solamente el nombre del fichero XML
        // Ejemplo: inventario_junio2026.xml
        String nombreFichero = ficheroXML.getName();

        // Quitamos la extensión .xml del nombre
        // Resultado: inventario_junio2026
        String nombreSinExtension = nombreFichero.substring(0, nombreFichero.lastIndexOf("."));

        // Obtenemos la fecha que aparece después del "_"
        // Ejemplo: inventario_junio2026 -> junio2026
        String fecha = nombreSinExtension.substring(nombreSinExtension.indexOf("_") + 1);

        // Variable donde vamos a guardar el beneficio total
        double beneficioTotal = 0;

        // Recorremos todos los productos de la lista
        for (ProductoEntity producto : productos) {

            // Sumamos el beneficio de cada producto al beneficio total
            // doubleValue() convierte el BigDecimal a double
            beneficioTotal += producto.getProfit().doubleValue();
        }

        // Creamos un objeto File que representa la carpeta donde
        // queremos guardar el fichero de resultado
        File carpeta = new File(path);

        // Comprobamos si la carpeta no existe
        if (!carpeta.exists()) {

            // Creamos la carpeta
            carpeta.mkdirs();
        }

        // Creamos el fichero de resultado dentro de la carpeta
        // Ejemplo: result_junio2026.txt
        File ficheroResultado = new File(carpeta, "result_" + fecha + ".txt");

        // Creamos un BufferedWriter para poder escribir texto
        // dentro del fichero de resultado
        BufferedWriter writer = new BufferedWriter(new FileWriter(ficheroResultado));

    }

    public void exportExcel(String path, String fileXml) throws JAXBException, IOException, ParseException {
        //TODO: Implementar
    }
}