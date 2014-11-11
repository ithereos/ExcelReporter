
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Random;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author 20102122399
 */
public class Inicio extends javax.swing.JFrame {
    
    private String rutaArchivo="ReporteExctel.xls";

    /**
     * Creates new form Inicio
     */
    public Inicio() {
        initComponents();
        Database.Open();
    }
    
    //----------- genera un archivo en EXcel
 public void CrearExcel()
 {
    File archivoXLS = new File(rutaArchivo);
    if(archivoXLS.exists()) archivoXLS.delete();
        try {
            archivoXLS.createNewFile();
            Workbook libro = new HSSFWorkbook(); //se crea el objeto 
            FileOutputStream archivo = new FileOutputStream(archivoXLS);
            Sheet hoja = libro.createSheet("Datos");// crea la hoja de Trabajo
            Random rnd = new Random();
            for(int f=0;f<100;f++){
                Row fila = hoja.createRow(f);
               for(int c=0;c<5;c++){
                  Cell celda = fila.createCell(c);
                  if(f==0){
                       celda.setCellValue("Encabezado #"+c);
                    }else{
                        celda.setCellValue(rnd.nextInt(999)+1); //genera un numero entre 1 y 100
                        }
                } //end for
           }//end For
            libro.write(archivo);
            archivo.close();
            //Para que lo abra en el programa predeterminado...
            Runtime.getRuntime().exec("cmd /c "+rutaArchivo); 
            System.out.println("Archivo "+rutaArchivo+" Creado");
        } catch (IOException ex) {
            System.err.println("Error creando el archivo "+ex.getMessage());
        }
 }   
 
 public void Oracle_Excel(){
     
    String nombres,apellidos,id,tdoc,genero,edad,historia;
    int i=1;
    int c=0;
    ResultSet reg = Database.consultar("SELECT * FROM PACIENTES");
    File archivoXLS = new File(rutaArchivo);
    if(archivoXLS.exists()) archivoXLS.delete();
        try {
            archivoXLS.createNewFile();
            Workbook libro = new HSSFWorkbook(); //se crea el objeto 
            FileOutputStream archivo = new FileOutputStream(archivoXLS);
            Sheet hoja = libro.createSheet("Datos");// crea la hoja de Trabajo
            Row fila = hoja.createRow(0);
            Cell celda = fila.createCell(0);
            celda.setCellValue("DNI");
            celda = fila.createCell(1);
            celda.setCellValue("TIPO");
            celda = fila.createCell(2);
            celda.setCellValue("NOMBRES");
            celda = fila.createCell(3);
            celda.setCellValue("APELLIDOS");
            celda = fila.createCell(4);
            celda.setCellValue("GENERO");
            celda = fila.createCell(5);
            celda.setCellValue("EDAD");
            celda = fila.createCell(6);
            celda.setCellValue("HISTORIACL");
            while (reg.next()){
                fila = hoja.createRow(i);
                id=reg.getString("DNI");
                tdoc=reg.getString("TIPODOCUMENTO");
                nombres = reg.getString("NOMBRE1")+ " "+reg.getString("NOMBRE2");
                apellidos = reg.getString("APELLIDO1")+" "+reg.getString("APELLIDO2");
                tdoc=reg.getString("TIPODOCUMENTO");
                genero = reg.getString("SEXO");
                edad = reg.getString("EDAD");
                historia = reg.getString("NUMERO_HISTORIA");
                celda = fila.createCell(0);
                celda.setCellValue(id);
                celda = fila.createCell(1);
                celda.setCellValue(tdoc);
                celda = fila.createCell(2);
                celda.setCellValue(nombres);
                celda = fila.createCell(3);
                celda.setCellValue(apellidos);
                celda = fila.createCell(4);
                celda.setCellValue(genero);
                celda = fila.createCell(5);
                celda.setCellValue(edad);
                celda = fila.createCell(6);
                celda.setCellValue(historia);
                i++;
                
                
            }
            libro.write(archivo);
            archivo.close();
            
         Runtime.getRuntime().exec("cmd /c "+rutaArchivo); 
            System.out.println("Archivo "+rutaArchivo+" Creado");
            
        } catch (SQLException ex) {
            System.err.println("Error creando el archivo "+ex.getMessage());
        }
        catch (IOException ex) {
            System.err.println("Error creando el archivo "+ex.getMessage());
        }
 }
 
 
 //---------- consulta de oracle a la consula
 public void Oracle_Consola()
{
    String nombres,apellidos,id,tdoc;
     ResultSet reg=Database.consultar("Select * from pacientes");
        try {
            int i=1;
            while(reg.next())
            {
               id=reg.getString("DNI");
               nombres = reg.getString("Nombre1")+ " "+reg.getString("Nombre2");
               tdoc=reg.getString("TIPODOCUMENTO");
               System.out.println("No. "+i+" Id  "+id+" Tipo "+tdoc+" Nombre: "+nombres);  
                i++;
            }//end while
        } catch (SQLException ex) {
              System.err.println("Se gener� el siguiente Error "+ex.getMessage());
        }
}

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        bttn_random = new javax.swing.JButton();
        bttn_export = new javax.swing.JButton();
        bttn_consult = new javax.swing.JButton();
        bttn_exit = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        bttn_random.setText("Random");
        bttn_random.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                bttn_randomActionPerformed(evt);
            }
        });

        bttn_export.setText("Exportar");
        bttn_export.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                bttn_exportActionPerformed(evt);
            }
        });

        bttn_consult.setText("Consultar");
        bttn_consult.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                bttn_consultActionPerformed(evt);
            }
        });

        bttn_exit.setText("Salir");

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(45, 45, 45)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(bttn_consult, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(bttn_random, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addGap(95, 95, 95)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(bttn_export, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(bttn_exit, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap(71, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(44, 44, 44)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(bttn_random)
                    .addComponent(bttn_export))
                .addGap(50, 50, 50)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(bttn_consult)
                    .addComponent(bttn_exit))
                .addContainerGap(56, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void bttn_randomActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_bttn_randomActionPerformed
        // TODO add your handling code here:
        CrearExcel();
    }//GEN-LAST:event_bttn_randomActionPerformed

    private void bttn_consultActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_bttn_consultActionPerformed
        // TODO add your handling code here:
        Oracle_Consola();
    }//GEN-LAST:event_bttn_consultActionPerformed

    private void bttn_exportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_bttn_exportActionPerformed
        // TODO add your handling code here:
        Oracle_Excel();
    }//GEN-LAST:event_bttn_exportActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Inicio.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Inicio.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Inicio.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Inicio.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Inicio().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton bttn_consult;
    private javax.swing.JButton bttn_exit;
    private javax.swing.JButton bttn_export;
    private javax.swing.JButton bttn_random;
    // End of variables declaration//GEN-END:variables
}
