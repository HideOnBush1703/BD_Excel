
package modelo;

import java.sql.Connection;
import java.sql.DriverManager;


public class Conexion {
    
    public static final String URL="jdbc:mysql://localhost:3306/tienda1?autoReconnet=true&useSSL=false";
    public static final String usuario="root";
    public static final String contraseña="alianzalima";
    
    public Connection getConnection(){
        Connection conexion=null;
        
        try{
            Class.forName("com.mysql.cj.jdbc.Driver");
            conexion= DriverManager.getConnection(URL,usuario,contraseña);
            //JOptionPane.showMessageDialog(null, "Conexion Exitosa");
            
        }catch(Exception ex){
            System.err.println("Error, "+ex);
        }
      return conexion;  
    }
    
}
