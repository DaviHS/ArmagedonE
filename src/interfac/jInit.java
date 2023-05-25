/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package interfac;

import javax.swing.ImageIcon;

/**
 * Tela Principal, reponsavel por interligar as outras pagina pela aba do menu
 * 
 */
public class jInit extends javax.swing.JFrame {

    public static boolean Excel;

    public static boolean Ficha;


    /**
     * Creates new form jPrincipal
     */
    public jInit() {
        initComponents();
    }


    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jTela = new javax.swing.JDesktopPane();
        jBarraInicial = new javax.swing.JMenuBar();
        jMenu1 = new javax.swing.JMenu();
        jExcel = new javax.swing.JMenuItem();
        jFicha = new javax.swing.JMenuItem();

        setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);
        setTitle("Projeto Armagedon");
        setIconImages(null);
        setResizable(false);
        addWindowListener(new java.awt.event.WindowAdapter() {
            public void windowActivated(java.awt.event.WindowEvent evt) {
                formWindowActivated(evt);
            }
            public void windowOpened(java.awt.event.WindowEvent evt) {
                formWindowOpened(evt);
            }
        });

        javax.swing.GroupLayout jTelaLayout = new javax.swing.GroupLayout(jTela);
        jTela.setLayout(jTelaLayout);
        jTelaLayout.setHorizontalGroup(
            jTelaLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 441, Short.MAX_VALUE)
        );
        jTelaLayout.setVerticalGroup(
            jTelaLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 257, Short.MAX_VALUE)
        );

        jMenu1.setText("Gerar");

        jExcel.setText("Excel (PF)");
        jExcel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jExcelActionPerformed(evt);
            }
        });
        jMenu1.add(jExcel);

        jFicha.setText("Ficha");
        jFicha.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jFichaActionPerformed(evt);
            }
        });
        jMenu1.add(jFicha);

        jBarraInicial.add(jMenu1);

        setJMenuBar(jBarraInicial);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 441, Short.MAX_VALUE)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addComponent(jTela))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 257, Short.MAX_VALUE)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addComponent(jTela))
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void formWindowOpened(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowOpened
        jInit.Excel = false;

        jInit.Ficha = false;

    }//GEN-LAST:event_formWindowOpened

    private void formWindowActivated(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowActivated
        // TODO add your handling code here:
        ImageIcon icon = new ImageIcon("src/imagens/armageddon (1).png");
        setIconImage(icon.getImage());
    }//GEN-LAST:event_formWindowActivated

    private void jExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jExcelActionPerformed
        // TODO add your handling code here:
                if (Excel == false) {
            jExcel obj = new jExcel();
            jTela.add(obj);

            obj.setVisible(true);

        }
        
    }//GEN-LAST:event_jExcelActionPerformed

    private void jFichaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jFichaActionPerformed
        // TODO add your handling code here:
            if (Ficha == false) {
            jFicha obj = new jFicha();
            jTela.add(obj);

            obj.setVisible(true);

        }
    }//GEN-LAST:event_jFichaActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {

        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(jInit.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>

        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(() -> {
            new jInit().setVisible(true);
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JMenuBar jBarraInicial;
    private javax.swing.JMenuItem jExcel;
    private javax.swing.JMenuItem jFicha;
    private javax.swing.JMenu jMenu1;
    public static javax.swing.JDesktopPane jTela;
    // End of variables declaration//GEN-END:variables

}
