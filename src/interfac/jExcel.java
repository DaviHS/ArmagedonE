/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JInternalFrame.java to edit this template
 */
package interfac;

import static excel.ExcelWriter.ExcelWriter;
import java.io.IOException;
import java.sql.SQLException;
import javax.swing.JOptionPane;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 * Usuario pode CONSULTAR, Cadastrar, Alterar e Excluir NOVO CLIENTE e/ou
 * TRANSPORTADORA
 *
 */
public class jExcel extends javax.swing.JInternalFrame {


    /**
     * Creates new form jifCadastro
     */
    public jExcel() {
        initComponents();
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jbExcel = new javax.swing.JButton();
        jtData2 = new com.toedter.calendar.JDateChooser();
        jtData1 = new com.toedter.calendar.JDateChooser();
        jlFilter = new javax.swing.JLabel();
        jlMes = new javax.swing.JLabel();
        jlAno = new javax.swing.JLabel();

        setClosable(true);
        setIconifiable(true);
        setTitle("Gerar Excel PF");
        addInternalFrameListener(new javax.swing.event.InternalFrameListener() {
            public void internalFrameActivated(javax.swing.event.InternalFrameEvent evt) {
            }
            public void internalFrameClosed(javax.swing.event.InternalFrameEvent evt) {
                formInternalFrameClosed(evt);
            }
            public void internalFrameClosing(javax.swing.event.InternalFrameEvent evt) {
            }
            public void internalFrameDeactivated(javax.swing.event.InternalFrameEvent evt) {
            }
            public void internalFrameDeiconified(javax.swing.event.InternalFrameEvent evt) {
            }
            public void internalFrameIconified(javax.swing.event.InternalFrameEvent evt) {
            }
            public void internalFrameOpened(javax.swing.event.InternalFrameEvent evt) {
                formInternalFrameOpened(evt);
            }
        });

        jbExcel.setText("Gerar PF");
        jbExcel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jbExcelActionPerformed(evt);
            }
        });

        jtData2.setDateFormatString("dd/MM/yyyy");

        jtData1.setDateFormatString("dd/MM/yyyy");

        jlFilter.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jlFilter.setText("Selecione o período:");

        jlMes.setText("De:");

        jlAno.setText("Até");

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jbExcel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jtData2, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jtData1, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jlFilter, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, 191, Short.MAX_VALUE)
                    .addGroup(javax.swing.GroupLayout.Alignment.LEADING, layout.createSequentialGroup()
                        .addGap(6, 6, 6)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jlAno, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.PREFERRED_SIZE, 108, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jlMes, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.PREFERRED_SIZE, 69, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jlFilter)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jlMes)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jtData1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jlAno)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jtData2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jbExcel, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(24, Short.MAX_VALUE))
        );

        setBounds(115, 20, 215, 224);
    }// </editor-fold>//GEN-END:initComponents

    private void formInternalFrameOpened(javax.swing.event.InternalFrameEvent evt) {//GEN-FIRST:event_formInternalFrameOpened
        // TODO add your handling code here:
        jInit.Excel = true;

    

    }//GEN-LAST:event_formInternalFrameOpened

    private void formInternalFrameClosed(javax.swing.event.InternalFrameEvent evt) {//GEN-FIRST:event_formInternalFrameClosed
        // TODO add your handling code here:
        jInit.Excel = false;
        dispose();
    }//GEN-LAST:event_formInternalFrameClosed

    private void jbExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jbExcelActionPerformed

        //Coleta as duas datas para verificar se os valores estão nulos... solicitando o preenchimento.
        Date dataVe1 = jtData1.getDate();
        Date dataVe2 = jtData2.getDate();

        if (dataVe1 == null && dataVe2 == null) {

            JOptionPane.showMessageDialog(null, "Preencher as datas solicitads para gerar o PF!");
            jtData1.requestFocus();

        } else if (dataVe1 == null) {

            JOptionPane.showMessageDialog(null, "Inserir data de partida!");
            jtData1.requestFocus();

        } else if (dataVe2 == null) {

            JOptionPane.showMessageDialog(null, "Inserir até qual data será realizada a busca.\nSe referente a só um dia, insira a mesma data.");
            jtData2.requestFocus();

        } else {

            try {

                //Formata a data para exibir o periodo da pesquisa no sistema
                SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yy");
                String data1Saida = dateFormat.format(jtData1.getDate());
                String data2Saida = dateFormat.format(jtData2.getDate());

                long hora = new Date().getHours();
                long minuto = new Date().getMinutes();

                System.out.println("Iniciando processo... " + hora + ":" + minuto + ". (" + data1Saida + " a " + data2Saida + ")");

                //Formata a data para o formato de pesquisa
                SimpleDateFormat dateFormatEntrada = new SimpleDateFormat("MM/dd/yyyy");
                String data1Entrada = dateFormatEntrada.format(jtData1.getDate());
                String data2Entrada = dateFormatEntrada.format(jtData2.getDate());

                ExcelWriter(data1Entrada, data2Entrada, dataVe1, dataVe2);

                JOptionPane.showMessageDialog(null, "Excel de Fundicao gerado com sucesso!");

            } catch (IOException | InvalidFormatException | SQLException err ) {
                JOptionPane.showMessageDialog(null, "Erro ao gerar aquivo PF" + err);
            } catch (ParseException ex) {
                java.util.logging.Logger.getLogger(jExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
            }

        }

    }//GEN-LAST:event_jbExcelActionPerformed


    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jbExcel;
    private javax.swing.JLabel jlAno;
    private javax.swing.JLabel jlFilter;
    private javax.swing.JLabel jlMes;
    private com.toedter.calendar.JDateChooser jtData1;
    private com.toedter.calendar.JDateChooser jtData2;
    // End of variables declaration//GEN-END:variables
}