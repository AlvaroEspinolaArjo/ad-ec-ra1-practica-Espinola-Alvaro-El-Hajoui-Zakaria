package org.educa.service;

import jakarta.xml.bind.JAXBException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.educa.dao.ProductoDAO;
import org.educa.entity.ProductoEntity;

import java.io.*;
import java.text.ParseException;
import java.util.List;

public class ProductoService {

    private ProductoDAO productoDAO = new ProductoDAO();

    /**
     * Lee los productos de un fichero XML.
     *
     * @param fileXml La ruta al fichero XML.
     * @return Una lista de productos leídos del fichero XML.
     * @throws JAXBException Excepción producida al procesar el fichero XML.
     */

    public List<ProductoEntity> readFile(String fileXml) throws JAXBException {
        return productoDAO.readFile(fileXml); // Llama al metodo readFile() de productoDAO y devuelve el resultado.
    }

    /**
     * Exporta un resumen de los productos a un fichero de texto.
     *
     * @param path La ruta de la carpeta donde se guardará el fichero.
     * @param fileXml La ruta al fichero XML.
     * @throws JAXBException Excepción producida al procesar el fichero XML.
     * @throws IOException Excepción producida al crear o escribir el fichero de texto.
     */

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

        // Creamos un BufferedWriter para poder escribir texto dentro del fichero de resultado
        BufferedWriter writer = new BufferedWriter(new FileWriter(ficheroResultado));

        // Escribimos la fecha en el fichero
        writer.write("Fecha: " + fecha);
        writer.newLine();

        // Escribimos el número total de productos
        writer.write("NumeroDeProductos: " + productos.size());
        writer.newLine();

        // Escribimos el beneficio total de todos los productos
        writer.write("BeneficioTotal: " + beneficioTotal);
        writer.newLine();

        // Escribimos la ruta absoluta del fichero XML
        writer.write("Ruta del fichero: " + ficheroXML.getAbsolutePath());
        writer.newLine();

        // Escribimos el nombre del fichero sin la extensión
        writer.write("Nombre del fichero: " + nombreSinExtension);
        writer.newLine();

        // Escribimos el tamaño del fichero XML en bytes
        writer.write("Tamaño del fichero: " + ficheroXML.length() + " bytes");

        // Cerramos el BufferedWriter para guardar correctamente toda la información y liberar el fichero
        writer.close();
    }

    /**
     * Exporta los productos de un fichero XML a un fichero Excel.
     * @param path La ruta de la carpeta donde se guardará el fichero Excel.
     * @param fileXml La ruta al fichero XML que contiene los productos.
     * @throws JAXBException Excepción producida al procesar el fichero XML.
     * @throws IOException Excepción producida al crear o escribir el fichero Excel.
     * @throws ParseException Excepción producida al procesar datos con formato incorrecto.
     */

    public void exportExcel(String path, String fileXml) throws JAXBException, IOException, ParseException {

        // Leemos el fichero XML y obtenemos la lista de productos
        List<ProductoEntity> productos = readFile(fileXml);

        // Creamos un objeto File que representa la carpeta donde guardaremos el Excel
        File carpeta = new File(path);

        // Si la carpeta no existe, la creamos
        if (!carpeta.exists()) {
            carpeta.mkdirs();
        }

        // Creamos un objeto File que representa el fichero XML
        File ficheroXML = new File(fileXml);

        // Obtenemos solamente el nombre del fichero XML
        String nombreFichero = ficheroXML.getName();

        // Quitamos la extensión .xml del nombre del fichero
        String nombreSinExtension = nombreFichero.substring(0, nombreFichero.lastIndexOf("."));

        // Obtenemos la parte que aparece después del "_"
        // Por ejemplo: inventario_junio2026.xml -> junio2026
        String fecha = nombreSinExtension.substring(nombreSinExtension.indexOf("_") + 1);

        // Creamos el nombre del fichero Excel
        // Por ejemplo: export_junio2026.xlsx
        File ficheroExcel = new File(carpeta,"export_" + fecha + ".xlsx");

        // Creamos un nuevo libro de Excel
        Workbook workbook = new XSSFWorkbook();

        // Creamos una hoja llamada "Productos"
        Sheet hoja = workbook.createSheet("Productos");

        // Creamos un array con los nombres de las columnas
        String[] columnas = {
                "Codigo",
                "Número de Serie",
                "Precio",
                "Descuento",
                "Precio Final",
                "Costes Envío",
                "Costes Almacenaje",
                "Beneficio"
        };

        // ESTILO DE LA CABECERA

        CellStyle estiloCabecera = workbook.createCellStyle();

        Font fuenteCabecera = workbook.createFont();
        fuenteCabecera.setBold(true);

        estiloCabecera.setFont(fuenteCabecera);
        estiloCabecera.setAlignment(HorizontalAlignment.CENTER);
        estiloCabecera.setVerticalAlignment(VerticalAlignment.CENTER);

        // Fondo blanco
        estiloCabecera.setFillForegroundColor(
                IndexedColors.WHITE.getIndex()
        );

        estiloCabecera.setFillPattern(
                FillPatternType.SOLID_FOREGROUND
        );

        // Bordes verdes
        estiloCabecera.setBorderTop(BorderStyle.THIN);
        estiloCabecera.setBorderBottom(BorderStyle.THIN);
        estiloCabecera.setBorderLeft(BorderStyle.THIN);
        estiloCabecera.setBorderRight(BorderStyle.THIN);

        estiloCabecera.setTopBorderColor(
                IndexedColors.GREEN.getIndex()
        );
        estiloCabecera.setBottomBorderColor(
                IndexedColors.GREEN.getIndex()
        );
        estiloCabecera.setLeftBorderColor(
                IndexedColors.GREEN.getIndex()
        );
        estiloCabecera.setRightBorderColor(
                IndexedColors.GREEN.getIndex()
        );

        // Creamos la primera fila de la hoja, que será la cabecera
        Row filaCabecera = hoja.createRow(0);

        for (int i = 0; i < columnas.length; i++) {

            Cell celda = filaCabecera.createCell(i);

            celda.setCellValue(columnas[i]);

            celda.setCellStyle(estiloCabecera);
        }

        // ESTILO FILA VERDE

        CellStyle estiloFilaVerde = workbook.createCellStyle();

        estiloFilaVerde.setFillForegroundColor(
                IndexedColors.LIGHT_GREEN.getIndex()
        );

        estiloFilaVerde.setFillPattern(
                FillPatternType.SOLID_FOREGROUND
        );

        estiloFilaVerde.setVerticalAlignment(
                VerticalAlignment.CENTER
        );

        estiloFilaVerde.setBorderTop(BorderStyle.THIN);
        estiloFilaVerde.setBorderBottom(BorderStyle.THIN);
        estiloFilaVerde.setBorderLeft(BorderStyle.THIN);
        estiloFilaVerde.setBorderRight(BorderStyle.THIN);

        estiloFilaVerde.setTopBorderColor(
                IndexedColors.GREEN.getIndex()
        );
        estiloFilaVerde.setBottomBorderColor(
                IndexedColors.GREEN.getIndex()
        );
        estiloFilaVerde.setLeftBorderColor(
                IndexedColors.GREEN.getIndex()
        );
        estiloFilaVerde.setRightBorderColor(
                IndexedColors.GREEN.getIndex()
        );

        // ESTILO FILA BLANCA

        CellStyle estiloFilaBlanca = workbook.createCellStyle();

        estiloFilaBlanca.setFillForegroundColor(
                IndexedColors.WHITE.getIndex()
        );

        estiloFilaBlanca.setFillPattern(
                FillPatternType.SOLID_FOREGROUND
        );

        estiloFilaBlanca.setVerticalAlignment(
                VerticalAlignment.CENTER
        );

        estiloFilaBlanca.setBorderTop(BorderStyle.THIN);
        estiloFilaBlanca.setBorderBottom(BorderStyle.THIN);
        estiloFilaBlanca.setBorderLeft(BorderStyle.THIN);
        estiloFilaBlanca.setBorderRight(BorderStyle.THIN);

        estiloFilaBlanca.setTopBorderColor(
                IndexedColors.GREEN.getIndex()
        );
        estiloFilaBlanca.setBottomBorderColor(
                IndexedColors.GREEN.getIndex()
        );
        estiloFilaBlanca.setLeftBorderColor(
                IndexedColors.GREEN.getIndex()
        );
        estiloFilaBlanca.setRightBorderColor(
                IndexedColors.GREEN.getIndex()
        );

        // FUENTE PARA LOS CÓDIGOS

        Font fuenteCodigo = workbook.createFont();
        fuenteCodigo.setBold(true);

        // ESTILO DE CÓDIGO VERDE

        CellStyle estiloCodigoVerde = workbook.createCellStyle();

        estiloCodigoVerde.cloneStyleFrom(estiloFilaVerde);
        estiloCodigoVerde.setFont(fuenteCodigo);

        // ESTILO DE CÓDIGO BLANCO

        CellStyle estiloCodigoBlanco = workbook.createCellStyle();

        estiloCodigoBlanco.cloneStyleFrom(estiloFilaBlanca);
        estiloCodigoBlanco.setFont(fuenteCodigo);
    }

}