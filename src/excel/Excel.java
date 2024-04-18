
package excel;


import java.io.File;
import java.sql.DriverManager;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import static javax.management.openmbean.SimpleType.STRING;
import java.sql.Connection;
import java.sql.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.logging.Level;
import java.util.logging.Logger;
import modelo.Conexion;
import org.apache.poi.ss.usermodel.DateUtil;

public class Excel {

    public static void main(String[] args) {
       
       //crearExcel(); 
       leerExcel();
       //cargarExcel();
       //cargarBD_Excel();
    }
    
    public static void crearExcel(){
        
        Workbook book=new XSSFWorkbook();
        Sheet sheet =  book.createSheet("Hola Java");
        
        Row row=sheet.createRow(0);//fila creada
        row.createCell(0).setCellValue("Hola Mundo");//llenando celdas con diferentes tipos de variables
        row.createCell(1).setCellValue(7.5);
        row.createCell(2).setCellValue(true);
        
        Cell celda=row.createCell(3);
        celda.setCellFormula(String.format("1+1", ""));
        
        Row rowUno=sheet.createRow(1);
        rowUno.createCell(0).setCellValue(7);//A2
        rowUno.createCell(1).setCellValue(8);//B2
        
        Cell celdados=rowUno.createCell(2);
        celdados.setCellFormula(String.format("A%d+B%d", 2,2));
        
        try {
            FileOutputStream archivo=new FileOutputStream("Excel.xlsx");
            book.write(archivo);
            archivo.close();  
        
        } catch (FileNotFoundException ex) {
            System.err.println("Error, "+ex);
        } catch (IOException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public static void leerExcel(){
        
        try{
            
            FileInputStream archivo = new FileInputStream(new File("C:\\Users\\Josue Pariguana\\Prueba1.xlsx"));
            XSSFWorkbook libroLectura = new XSSFWorkbook(archivo);
            XSSFSheet hojaLectura = libroLectura.getSheetAt(0); //detectando el excel a leer.

            int numFilas = hojaLectura.getLastRowNum();//hallando el numero de filas en la hoja

            for (int i = 0; i <= numFilas; i++) { //obeteniendo todos los valores de cada fila
                Row fila = hojaLectura.getRow(i);
                int numCol = fila.getLastCellNum();//hallando el numero de columnas en la hoja

                for (int j = 0; j < numCol; j++) {
                    Cell celda = fila.getCell(j);
     
                    switch (celda.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:
                            if (DateUtil.isCellDateFormatted(celda)) {
                                // Si la celda contiene una fecha
                                System.out.print(celda.getDateCellValue() + " ");
                            } else {
                                // Si la celda contiene un valor numérico
                                double valorNumerico = celda.getNumericCellValue();
                                if (valorNumerico == (int) valorNumerico) {
                                    // Si el valor es un entero, imprímelo como un entero
                                    System.out.print((int) valorNumerico + " ");
                                } else {
                                    // Si el valor tiene decimales, imprímelo como un double
                                    System.out.print(valorNumerico + " ");
                                }
                            }
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(celda.getStringCellValue() + " ");
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            System.out.print(celda.getCellFormula() + " ");
                            break;
                    }
                }
                System.out.println("");
            }
            
        }catch(Exception ex){
            System.err.println("Error, "+ex);
        }
        
    }
    
    public static void cargarExcel(){
        
        Conexion con=new Conexion();
        PreparedStatement ps=null;
        
        try{
            
            Connection conexion=con.getConnection();
            FileInputStream archivo = new FileInputStream(new File("C:\\Users\\Josue Pariguana\\Prueba1.xlsx"));
            XSSFWorkbook libroLectura = new XSSFWorkbook(archivo);
            XSSFSheet hojaLectura = libroLectura.getSheetAt(0);
            
            int numFilas=hojaLectura.getLastRowNum();
            
            for(int i=1;i<=numFilas; i++){
                
                Row fila=hojaLectura.getRow(i);//una vez que hemos obtenido todos los datos de la fila.
              
                java.util.Date utilDate = fila.getCell(3).getDateCellValue();
                java.sql.Date sqlDate = new java.sql.Date(utilDate.getTime());

                ps = conexion.prepareStatement("insert into producto (idproducto, nombre, precio, fecha_venta, idcategoria, cantidad) values (?,?,?,?,?,?)");
                ps.setInt(1, (int) fila.getCell(0).getNumericCellValue());
                ps.setString(2, fila.getCell(1).getStringCellValue());
                ps.setDouble(3, fila.getCell(2).getNumericCellValue());
                ps.setDate(4, sqlDate);
                ps.setInt(5, (int) fila.getCell(4).getNumericCellValue());
                ps.setInt(6, (int) fila.getCell(5).getNumericCellValue());
                ps.executeUpdate();
            }
            
            conexion.close();
            
        }catch(Exception ex){
            System.err.println("Error, "+ex);
        }
    }
    
    public static void cargarBD_Excel(){
        //NUEVO EXCEL OJO
        Workbook libro=new XSSFWorkbook();
        Sheet hoja =  libro.createSheet("Reporte Productos");
        Conexion con=new Conexion();
        PreparedStatement ps=null;
        ResultSet rs=null;
        
        String[] cabeceras=new String[]{"IdProducto","Nombre","Precio","Fecha Venta","IdCategoría","Cantidad"};
        
        Row filaCabeceras=hoja.createRow(0);//Fila cabeceras de las columnas(esto no lo puedo traer de la BD)
        for(int i=0;i<cabeceras.length;i++){
            Cell celda=filaCabeceras.createCell(i);
            celda.setCellValue(cabeceras[i]);
        }
        
        int numFila=1; //fila 0 ya llena por las cabeceras
        
        try {
            
            Connection conexion=con.getConnection();
            
            ps=conexion.prepareStatement("select idproducto,nombre,precio,fecha_venta,idCategoria,cantidad from producto");
            rs=ps.executeQuery();
            
            int numCol=rs.getMetaData().getColumnCount();//numero de columas de la consulta hecha
            while(rs.next()){
                Row filaDatos=hoja.createRow(numFila);
                
                for(int i=0;i<numCol;i++){
                    Cell celda=filaDatos.createCell(i);
                    
                    if(i==0 || i==2 || i==4 || i==5){
                        celda.setCellValue(rs.getDouble(i+1));
                    }
                    else if(i==1){
                        celda.setCellValue(rs.getString(i+1));
                    }
                    else if(i==3){
                        celda.setCellValue(rs.getDate(i+1));
                    }
                }
                numFila++;
            }
            
            //esto va al final, antes sale error... -->
            FileOutputStream archivo=new FileOutputStream("ReporteProductos.xlsx");
            libro.write(archivo);
            archivo.close();  
        
        } catch (FileNotFoundException ex) {
            System.err.println("Error, "+ex);
        } catch (IOException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SQLException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
}
