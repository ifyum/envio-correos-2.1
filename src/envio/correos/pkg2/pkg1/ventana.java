/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package envio.correos.pkg2.pkg1;
import java.awt.Dimension;
import java.io.File;
import java.io.IOException;
import static java.lang.Integer.parseInt;
import java.security.Principal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableRowSorter;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 *
 * @author Ruben Barrera
 */
public class ventana extends javax.swing.JFrame {
 private static DefaultTableModel modelo;
    private TableRowSorter trsFiltro;

    //para exportar
    private JFileChooser FileChooser = new JFileChooser();
    private Vector columna = new Vector();
    private Vector filas = new Vector();
    private static int tabla_ancho = 0;
    private static int tabla_alto = 0;
    int fecha,fechasuma=0,horai,horas,dia1=1;
    int contadorcelda=0,contador=1;  
    Date date = new Date();
    Date Dia=new Date();
     String cambiodia,nombre,error2="";
    
      public void CrearTabla(File file) throws IOException {

DateFormat dateFormat = new SimpleDateFormat("MM/yyyy");
DateFormat datefecha = new SimpleDateFormat("dd");
String convertir;

convertir=datefecha.format(Dia).toString();
fecha=parseInt(convertir);
fechasuma=fecha;
System.out.println(convertir);

//dia de la semana
  String[] dias={"Domingo","Lunes","Martes", "Miércoles","Jueves","Viernes","Sábado"};
        Date hoy=new Date();
      int numeroDia=0;
      Calendar cal= Calendar.getInstance();
      cal.setTime(hoy);
      numeroDia=cal.get(Calendar.DAY_OF_WEEK);
      System.out.println("hoy es "+ dias[numeroDia - 1]);
cambiodia=dias[numeroDia - 1];
/// fin dia de semana

Workbook workbook = null;



        try {
            workbook = Workbook.getWorkbook(file);

            Sheet sheet = workbook.getSheet(0);
            columna.clear();

            for (int i = 0; i < sheet.getColumns(); i++) {
                Cell cell1 = sheet.getCell(i, 0);
                columna.add(cell1.getContents());
                
            }
            filas.clear();

            for (int j = 1; j < sheet.getRows(); j++) {
  Vector d = new Vector();
             
                for (int i = 0; i < sheet.getColumns(); i++) {
     Cell celda2 = sheet.getCell(0, j);
     //nombre fila 1               
     nombre=celda2.getContents();
                    //lista completa
                     Cell cell = sheet.getCell(i, j);
         
                     if(i==0 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
                      if( i==1 || i==4|| i==7|| i==10|| i==13 || i==16 || i==19){
       contador=contador+1;
         System.out.println("fila: "+i+" columna: "+j+"= "+fechasuma+"/"+dateFormat.format(date) );
  

 d.add(fechasuma+"/"+dateFormat.format(date));
 
contadorcelda=contadorcelda+1;
fechasuma=fechasuma+1;

   if( contador>7){
       contador=1;
       fechasuma=fecha;
       System.out.println("reset");
   }

   
   
   
   
   
   
   } 
          if(i==2 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
        if(i==3 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
     
 
   
    if(i==5 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
   
    if(i==6 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
     if(i==8 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
   
    if(i==9 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
   
    if(i==11 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
    
     if(i==12 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
    if(i==14 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
    
     if(i==15 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
     
        if(i==17 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
    
     if(i==18 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
     
     if(i==20 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
    
     if(i==21 && j < sheet.getRows()){
                      d.add(cell.getContents());
                     }
     
   //fin fehca
   
   
      
       
     
           
                }
                d.add("\n");
                //filas.add(d);
                modelo.addRow(d);
             
       
                
            }

        } catch (BiffException e) {
            e.printStackTrace();
        }
    }
    /**
     * Creates new form ventana
     */
    public ventana() {
        initComponents();
            modelo = new DefaultTableModel();
            
          String[] dias={"Domingo","Lunes","Martes", "Miércoles","Jueves","Viernes","Sábado"};
        Date hoy=new Date();
      int numeroDia=0;
      Calendar cal= Calendar.getInstance();
      cal.setTime(hoy);
      numeroDia=cal.get(Calendar.DAY_OF_WEEK);
      int dia2=numeroDia-dia1;
      System.out.println("hoy es "+ dias[dia2]);   
           System.out.println("hoy es "+ dia2);   
          
      
        modelo.addColumn("NOMBRE");
        for(int i=1;i<=7;i=i+1){
            
          
        modelo.addColumn(""+dias[dia2]);
        modelo.addColumn("HORA INGRESO");
        modelo.addColumn("HORA SALIDA");
        dia2=dia2+1; 
if(dia2==7 ){
     dia2=0;
            }

        /*   
        if(dia1==0 ){
     dia1=7;
            } */
        }
       
        
      
        
        this.tablaListado.setModel(modelo);
   
      
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jScrollPane2 = new javax.swing.JScrollPane();
        tablaListado = new javax.swing.JTable();
        jMenuBar2 = new javax.swing.JMenuBar();
        jMenu4 = new javax.swing.JMenu();
        jMenuItem1 = new javax.swing.JMenuItem();
        jMenuItem2 = new javax.swing.JMenuItem();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        tablaListado.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jScrollPane2.setViewportView(tablaListado);

        jMenu4.setText("Archivo");

        jMenuItem1.setAccelerator(javax.swing.KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_F1, 0));
        jMenuItem1.setText("Cargar Excel");
        jMenuItem1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem1ActionPerformed(evt);
            }
        });
        jMenu4.add(jMenuItem1);

        jMenuItem2.setAccelerator(javax.swing.KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_F2, 0));
        jMenuItem2.setText("Exportar Excel");
        jMenuItem2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem2ActionPerformed(evt);
            }
        });
        jMenu4.add(jMenuItem2);

        jMenuBar2.add(jMenu4);

        setJMenuBar(jMenuBar2);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 631, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 445, Short.MAX_VALUE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jMenuItem1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem1ActionPerformed
        //importar y enviar
        // TODO add your handling code here:
        FileChooser.showDialog(null, "Importar Hoja ");
        File file = FileChooser.getSelectedFile();
        if (!file.getName().endsWith("xls")) {

            JOptionPane.showMessageDialog(null, "Seleccione un archivo excel...", "Error", JOptionPane.ERROR_MESSAGE);
        } else {
            try {

                CrearTabla(file); //a nuestro metodo CrearTabla le enviamos el file

            } catch (IOException ex) {
                Logger.getLogger(ventana.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }//GEN-LAST:event_jMenuItem1ActionPerformed

    private void jMenuItem2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem2ActionPerformed
        JFileChooser dialog = new JFileChooser();
        int opcion = dialog.showSaveDialog(ventana.this);

        if (opcion == JFileChooser.APPROVE_OPTION) {

            File dir = dialog.getSelectedFile();

            try {
                List<JTable> tb = new ArrayList<JTable>();
                tb.add(tablaListado);

                //-------------------
                export_excel excelExporter = new export_excel(tb, new File(dir.getAbsolutePath() + ".xls"));
                if (excelExporter.export()) {
                    JOptionPane.showMessageDialog(null, "TABLAS EXPORTADOS CON EXITOS!");
                }
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
    }//GEN-LAST:event_jMenuItem2ActionPerformed

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
            java.util.logging.Logger.getLogger(ventana.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ventana.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ventana.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ventana.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ventana().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JMenu jMenu4;
    private javax.swing.JMenuBar jMenuBar2;
    private javax.swing.JMenuItem jMenuItem1;
    private javax.swing.JMenuItem jMenuItem2;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JTable tablaListado;
    // End of variables declaration//GEN-END:variables
}
