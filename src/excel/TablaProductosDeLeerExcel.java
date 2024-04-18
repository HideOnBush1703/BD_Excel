
package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TablaProductosDeLeerExcel extends javax.swing.JFrame {

    public TablaProductosDeLeerExcel() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jScrollPane1 = new javax.swing.JScrollPane();
        TablaExcel = new javax.swing.JTable();
        BotonCargar = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        TablaExcel.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null}
            },
            new String [] {
                "IdProducto", "Nombre", "Precio", "Fecha Venta", "IdCategoria", "Cantidad"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.Integer.class, java.lang.String.class, java.lang.Double.class, java.lang.String.class, java.lang.Integer.class, java.lang.Integer.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false, false
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        jScrollPane1.setViewportView(TablaExcel);

        BotonCargar.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        BotonCargar.setText("Cargar Tabla");
        BotonCargar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                BotonCargarActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(231, 231, 231)
                        .addComponent(BotonCargar))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(23, 23, 23)
                        .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 593, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap(29, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(28, 28, 28)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 359, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(38, 38, 38)
                .addComponent(BotonCargar)
                .addContainerGap(42, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void BotonCargarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_BotonCargarActionPerformed
      
        DefaultTableModel modeloTabla = new DefaultTableModel();
        TablaExcel.setModel(modeloTabla);
        
        try {
            
            modeloTabla.addColumn("IdProducto");
            modeloTabla.addColumn("Nombre");
            modeloTabla.addColumn("Precio");
            modeloTabla.addColumn("Fecha Venta");
            modeloTabla.addColumn("IdCategoria");
            modeloTabla.addColumn("Cantidad");
            
            FileInputStream archivo = new FileInputStream(new File("C:\\Users\\Josue Pariguana\\Prueba1.xlsx"));
            XSSFWorkbook libroLectura = new XSSFWorkbook(archivo);
            XSSFSheet hojaLectura = libroLectura.getSheetAt(0);
            
            int numFilas = hojaLectura.getLastRowNum();//hallando el numero de filas
            for (int i = 1; i <= numFilas; i++) { //obeteniendo todos los valores de cada fila
                Row fila = hojaLectura.getRow(i);
                int numCol = fila.getLastCellNum();//hallando el numero de columnas en la hoja

                Object arregloObjetos[] = new Object[numCol];
                
                for (int j = 0; j < numCol; j++) {
                    Cell celda = fila.getCell(j);
                    
                    switch (celda.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:
                            if (DateUtil.isCellDateFormatted(celda)) {
                                // Si la celda contiene una fecha
                                arregloObjetos[j]=celda.getDateCellValue(); //llenamos el arreglo
                            } else {
                                // Si la celda contiene un valor numérico
                                double valorNumerico = celda.getNumericCellValue();
                                if (valorNumerico == (int) valorNumerico) {
                                    // Si el valor es un entero, imprímelo como un entero
                                    arregloObjetos[j]=(int) valorNumerico; //llenamos el arreglo
                                } else {
                                    // Si el valor tiene decimales, imprímelo como un double
                                    arregloObjetos[j]=valorNumerico; //llenamos el arreglo
                                }
                            }
                            break;
                        case Cell.CELL_TYPE_STRING:
                            arregloObjetos[j]=celda.getStringCellValue(); //llenamos el arreglo
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            arregloObjetos[j]=celda.getCellFormula(); //llenamos el arreglo
                            break;
                    }
                    
                }
                modeloTabla.addRow(arregloObjetos);
            }
            
        
        } catch (Exception ex) {
            System.err.println("Error, "+ex);;
        }
        
        
    }//GEN-LAST:event_BotonCargarActionPerformed

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
            java.util.logging.Logger.getLogger(TablaProductosDeLeerExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(TablaProductosDeLeerExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(TablaProductosDeLeerExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(TablaProductosDeLeerExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new TablaProductosDeLeerExcel().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton BotonCargar;
    private javax.swing.JTable TablaExcel;
    private javax.swing.JScrollPane jScrollPane1;
    // End of variables declaration//GEN-END:variables
}
