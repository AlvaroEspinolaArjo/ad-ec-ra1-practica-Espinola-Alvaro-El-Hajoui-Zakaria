package org.educa.service;

import jakarta.xml.bind.JAXBException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.educa.dao.ProductoDAO;
import org.educa.entity.ProductoEntity;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.FileOutputStream;
import java.io.IOException;
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
        return productoDAO.readFile(fileXml);
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

        List<ProductoEntity> productos = readFile(fileXml);

        File ficheroXML = new File(fileXml);

        String nombreFichero = ficheroXML.getName();


        String nombreSinExtension = nombreFichero.substring(
                0, nombreFichero.lastIndexOf(".")
        );

        String fecha = nombreSinExtension.substring(
                nombreSinExtension.indexOf("_") + 1
        );

        double beneficioTotal = 0;

        for (ProductoEntity producto : productos) {
            beneficioTotal += producto.getProfit().doubleValue();
        }

        File carpeta = new File(path);

        if (!carpeta.exists()) {
            carpeta.mkdirs();
        }

        File ficheroResultado = new File(
                carpeta,
                "result_" + fecha + ".txt"
        );

        BufferedWriter writer = new BufferedWriter(
                new FileWriter(ficheroResultado)
        );

        writer.write("Fecha: " + fecha);
        writer.newLine();

        writer.write("NumeroDeProductos: " + productos.size());
        writer.newLine();

        writer.write("BeneficioTotal: " + beneficioTotal);
        writer.newLine();

        writer.write("Ruta del fichero: " + ficheroXML.getAbsolutePath());
        writer.newLine();

        writer.write("Nombre del fichero: " + nombreSinExtension);
        writer.newLine();

        writer.write("Tamaño del fichero: " + ficheroXML.length() + " bytes");

        writer.close();
    }

    /**
     * Exporta los productos de un fichero XML a un fichero Excel.
     *
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
        String nombreSinExtension = nombreFichero.substring(
                0, nombreFichero.lastIndexOf(".")
        );

        // Obtenemos la parte que aparece después del "_"
        // Por ejemplo: inventario_junio2026.xml -> junio2026
        String fecha = nombreSinExtension.substring(
                nombreSinExtension.indexOf("_") + 1
        );

        // Creamos el nombre del fichero Excel
        // Por ejemplo: export_junio2026.xlsx
        File ficheroExcel = new File(
                carpeta,
                "export_" + fecha + ".xlsx"
        );

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

        // FORMATO DE EUROS PARA FILA VERDE

        CellStyle estiloEurosVerde = workbook.createCellStyle();

        estiloEurosVerde.cloneStyleFrom(estiloFilaVerde);

        estiloEurosVerde.setDataFormat(
                workbook.createDataFormat().getFormat("#,##0.00 €")
        );

        // FORMATO DE EUROS PARA FILA BLANCA


        CellStyle estiloEurosBlanco = workbook.createCellStyle();

        estiloEurosBlanco.cloneStyleFrom(estiloFilaBlanca);

        estiloEurosBlanco.setDataFormat(
                workbook.createDataFormat().getFormat("#,##0.00 €")
        );

        // FORMATO DE PORCENTAJE PARA FILA VERDE


        CellStyle estiloPorcentajeVerde = workbook.createCellStyle();

        estiloPorcentajeVerde.cloneStyleFrom(estiloFilaVerde);

        estiloPorcentajeVerde.setDataFormat(
                workbook.createDataFormat().getFormat("0.00%")
        );

        // FORMATO DE PORCENTAJE PARA FILA BLANCA


        CellStyle estiloPorcentajeBlanco = workbook.createCellStyle();

        estiloPorcentajeBlanco.cloneStyleFrom(estiloFilaBlanca);

        estiloPorcentajeBlanco.setDataFormat(
                workbook.createDataFormat().getFormat("0.00%")
        );

        // RECORREMOS TODOS LOS PRODUCTOS


        for (int i = 0; i < productos.size(); i++) {

            // Obtenemos el producto que corresponde a la posición actual
            ProductoEntity productoEntity = productos.get(i);

            // Creamos una fila nueva en Excel
            // Sumamos 1 porque la fila 0 es la cabecera
            Row fila = hoja.createRow(i + 1);

            // Elegimos el color de la fila
            CellStyle estiloFila;
            CellStyle estiloCodigo;
            CellStyle estiloEuros;
            CellStyle estiloPorcentaje;

            if (i % 2 == 0) {
                estiloFila = estiloFilaVerde;
                estiloCodigo = estiloCodigoVerde;
                estiloEuros = estiloEurosVerde;
                estiloPorcentaje = estiloPorcentajeVerde;
            } else {
                estiloFila = estiloFilaBlanca;
                estiloCodigo = estiloCodigoBlanco;
                estiloEuros = estiloEurosBlanco;
                estiloPorcentaje = estiloPorcentajeBlanco;
            }

            // CÓDIGO


            Cell codigo = fila.createCell(0);

            codigo.setCellValue(
                    productoEntity.getProducto().getCodigo()
            );

            // Aplicamos color y negrita
            codigo.setCellStyle(estiloCodigo);

            // NÚMERO DE SERIE


            Cell numeroSerie = fila.createCell(1);

            numeroSerie.setCellValue(
                    productoEntity.getProducto().getNumeroSerie()
            );

            numeroSerie.setCellStyle(estiloFila);

            // PRECIO

            Cell precio = fila.createCell(2);

            precio.setCellValue(
                    productoEntity.getProducto().getPrecio().doubleValue()
            );

            precio.setCellStyle(estiloEuros);

            // DESCUENTO


            Cell descuento = fila.createCell(3);

            descuento.setCellValue(
                    productoEntity.getProducto().getDescuento().doubleValue() / 100
            );

            descuento.setCellStyle(estiloPorcentaje);


            // PRECIO FINAL


            Cell precioFinal = fila.createCell(4);

            precioFinal.setCellValue(
                    productoEntity.getPrecioFinal().doubleValue()
            );

            precioFinal.setCellStyle(estiloEuros);


            // COSTES DE ENVÍO


            Cell costesEnvio = fila.createCell(5);

            costesEnvio.setCellValue(
                    productoEntity.getProducto()
                            .getCostes()
                            .getCostesEnvio()
                            .doubleValue()
            );

            costesEnvio.setCellStyle(estiloEuros);

            // COSTES DE ALMACENAJE


            Cell costesAlmacenaje = fila.createCell(6);

            costesAlmacenaje.setCellValue(
                    productoEntity.getProducto()
                            .getCostes()
                            .getCostesAlmacenaje()
                            .doubleValue()
            );

            costesAlmacenaje.setCellStyle(estiloEuros);


            // BENEFICIO


            Cell beneficio = fila.createCell(7);

            beneficio.setCellValue(
                    productoEntity.getProfit().doubleValue()
            );

            beneficio.setCellStyle(estiloEuros);
        }


        // AJUSTAMOS EL ANCHO DE LAS COLUMNAS


        for (int i = 0; i < columnas.length; i++) {
            hoja.autoSizeColumn(i);
        }

        // Damos espacio al número de serie
        hoja.setColumnWidth(1, 20 * 256);

        // GUARDAMOS EL EXCEL

        FileOutputStream salida = new FileOutputStream(ficheroExcel);

        // Escribimos el contenido del Workbook en el fichero Excel
        workbook.write(salida);

        // Cerramos el flujo de salida
        salida.close();

        // Cerramos el Workbook
        workbook.close();
    }
}