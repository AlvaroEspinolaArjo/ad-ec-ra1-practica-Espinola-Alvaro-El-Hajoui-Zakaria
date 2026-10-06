package org.educa.dao;

import generated.Producto;
import generated.Productos;
import jakarta.xml.bind.JAXBContext;
import jakarta.xml.bind.JAXBException;
import jakarta.xml.bind.Unmarshaller;
import org.educa.entity.ProductoEntity;

import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

public class ProductoDAO {

    // Lee el archivo XML y devuelve una lista de productos
    public List<ProductoEntity> readFile(String fileXml) throws JAXBException {

        // Creamos una lista
        List<ProductoEntity> lista = new ArrayList<>();

        // Creamos el contexto JAXB indicando la clase que representa al XML
        JAXBContext context = JAXBContext.newInstance(Productos.class);

        // Creamos el objeto que se encarga de leer el XML
        Unmarshaller unmarshaller = context.createUnmarshaller();

        // Buscamos el archivo XML dentro de la carpeta resources
        InputStream inputStream = getClass().getClassLoader().getResourceAsStream("xml/inventario_junio2026.xml");

        // Convertimos el contenido del XML en un objeto Productos
        Productos productos = (Productos) unmarshaller.unmarshal(inputStream);

        // Recorremos todos los productos que hay en el XML
        for (Producto producto : productos.getProducto()) {

            // Creamos un objeto ProductoEntity para cada producto
            ProductoEntity entity = new ProductoEntity();

            // Guardamos el producto dentro de ProductoEntity
            entity.setProducto(producto);

            // Obtenemos el precio y el descuento del producto
            BigDecimal precio = producto.getPrecio();
            BigDecimal descuento = producto.getDescuento();

            // Calculamos el precio final aplicando el descuento
            BigDecimal precioFinal = precio.subtract(precio.multiply(descuento).divide(new BigDecimal("100")));

            // Guardamos el precio final
            entity.setPrecioFinal(precioFinal);

            // Obtenemos los costes de envío y almacenaje
            BigDecimal coste = producto.getCostes().getCostesEnvio().add(producto.getCostes().getCostesAlmacenaje());

            // Guardamos el coste total
            entity.setCost(coste);

            // Calculamos el beneficio restando los costes al precio final
            entity.setProfit(precioFinal.subtract(coste));

            // Añadimos el producto a la lista
            lista.add(entity);
        }

        // Devolvemos la lista con todos los productos
        return lista;
    }
}