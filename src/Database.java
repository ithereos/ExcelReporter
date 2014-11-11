import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.logging.Level;
import java.util.logging.Logger;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author user23
 */
public class Database {
    public static  String url="jdbc:oracle:thin:@127.0.0.1:1521:xe";
    public static  String user="system";
    public static  String password="123456";
    public static Connection connection;
    
    public static void Open()
    {
        try {
                Class.forName("oracle.jdbc.driver.OracleDriver");
                connection=DriverManager.getConnection(url,user,password);
            } catch (ClassNotFoundException e) {
                e.printStackTrace();
            } catch (SQLException e) {
                //Imprime el mensaje de la exception lanzada en pl/sql si el valor es diferente de 1
                e.printStackTrace();
            }
    }//end Open
//-- consultas de tipo INSERT, DELETE, UPDATE    
    public static boolean  Execute(String SQL)
    {
        try { 
            Statement sentencia; 
            sentencia = connection.createStatement(ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY); 
            sentencia.executeUpdate(SQL); 
           connection.commit(); 
            sentencia.close(); 
             
        } catch (SQLException e) { 
            e.printStackTrace(); 
            System.err.println("Error SQL  "+e.getMessage());
            return false; 
        }         
        return true; 
    }//end excecute 
        
          
   // CIERA LA CONEXION A BD
    public static void Close()
    {
        try {
            connection.close();
        } catch (SQLException ex) {
            System.err.println("Error "+ex.getMessage());
        }
    }
 //Ejecuta una consulta de tipo select
    public static ResultSet consultar(String sql) { 
        ResultSet resultado = null; 
        try { 
            Statement sentencia; 
            sentencia = connection.createStatement(ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY); 
            resultado = sentencia.executeQuery(sql); 
        } catch (SQLException e) { 
            e.printStackTrace(); 
            return null; 
        } 
        return resultado; 
    } 
    
    
    
    
}//end class