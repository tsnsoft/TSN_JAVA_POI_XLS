package tsn_java_poi.msexcel;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ReceiptExcel extends javax.swing.JFrame {
    private static final long serialVersionUID = 1L;

    class TThread1 extends Thread { // Поток запуска MS Excel

        public void run() {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator"); // Текущий катаолог
            try {
                modifData(dir + "receipt_template.xls", dir + "receipt.xls", jTextField_FIO.getText(),
                        jTextField_Adress.getText(), jTextField_SumPL.getText(),
                        jTextField_SummUS.getText()); // Вызов метода создания отчета
                Desktop.getDesktop().open(new File(dir + "receipt.xls")); // Запуск отчета в MS Excel
            } catch (Exception ex) {
                System.err.println("Error modifData!");
                ex.printStackTrace();
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
    }

    Thread tThread1; // Переменная-поток

    private void modifData(String inputFileName, String outputFileName, String FIO, String Adress,
            String SummPL, String SummUS) throws IOException {
        // Метод создания отчета
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(inputFileName)); // Файл-шаблон MS Excel
        HSSFWorkbook wb = new HSSFWorkbook(fs); // Документ MS Excel
        HSSFSheet sheet = wb.getSheetAt(0); // Первый лист в документе MS Excel

        sheet.getRow(12).getCell(3).setCellValue(FIO);
        sheet.getRow(13).getCell(3).setCellValue(Adress);
        sheet.getRow(14).getCell(3).setCellValue(SummPL);
        sheet.getRow(14).getCell(8).setCellValue(SummUS);

        int sum = Integer.parseInt(SummPL) + Integer.parseInt(SummUS);
        sheet.getRow(15).getCell(3).setCellValue(sum);
        
        try (FileOutputStream fileOut = new FileOutputStream(outputFileName)) {
            wb.write(fileOut);
        }
    }

    public ReceiptExcel() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jButton1 = new javax.swing.JButton();
        jTextField_FIO = new javax.swing.JTextField();
        jTextField_Adress = new javax.swing.JTextField();
        jTextField_SumPL = new javax.swing.JTextField();
        jTextField_SummUS = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Квитанция в Excel");
        setResizable(false);
        getContentPane().setLayout(null);

        jButton1.setText("в Excel");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        getContentPane().add(jButton1);
        jButton1.setBounds(620, 370, 71, 27);

        jTextField_FIO.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        getContentPane().add(jTextField_FIO);
        jTextField_FIO.setBounds(150, 247, 550, 27);

        jTextField_Adress.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        jTextField_Adress.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        getContentPane().add(jTextField_Adress);
        jTextField_Adress.setBounds(150, 270, 550, 27);

        jTextField_SumPL.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        jTextField_SumPL.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        getContentPane().add(jTextField_SumPL);
        jTextField_SumPL.setBounds(150, 295, 50, 27);

        jTextField_SummUS.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        jTextField_SummUS.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        getContentPane().add(jTextField_SummUS);
        jTextField_SummUS.setBounds(460, 295, 60, 27);

        jLabel2.setIcon(new javax.swing.ImageIcon(getClass().getResource("/tsn_java_poi/msexcel/receipt.png"))); // NOI18N
        jLabel2.setText("jLabel2");
        getContentPane().add(jLabel2);
        jLabel2.setBounds(0, 0, 710, 410);

        setSize(new java.awt.Dimension(722, 436));
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new TThread1().start();
    }//GEN-LAST:event_jButton1ActionPerformed

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
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ReceiptExcel().setVisible(true);

            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JTextField jTextField_Adress;
    private javax.swing.JTextField jTextField_FIO;
    private javax.swing.JTextField jTextField_SumPL;
    private javax.swing.JTextField jTextField_SummUS;
    // End of variables declaration//GEN-END:variables
}
