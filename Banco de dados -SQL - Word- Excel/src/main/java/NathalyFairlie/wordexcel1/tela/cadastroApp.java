package NathalyFairlie.wordexcel1.tela;

import NathalyFairlie.wordexcel1.DAO.PessoaDAO;
import NathalyFairlie.wordexcel1.model.Pessoa;
import bNathalyFairlie.wordexcel1.office.ExportarExcel;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.util.List;
import javax.swing.JOptionPane;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class cadastroApp extends javax.swing.JFrame {

    private PessoaDAO pessoaDAO;

    public cadastroApp() {
        initComponents();
        pessoaDAO = new PessoaDAO();
        atualizarComboPessoas();
        btnSalvar.setVisible(false); // Inicialmente, o botão "Salvar" é invisível
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        btnExcel = new javax.swing.JButton();
        btnInserir = new javax.swing.JButton();
        lblNome = new javax.swing.JLabel();
        txtNome = new javax.swing.JTextField();
        txtIdade = new javax.swing.JTextField();
        lblIdade = new javax.swing.JLabel();
        txtProfissao = new javax.swing.JTextField();
        lblProfissao = new javax.swing.JLabel();
        btnWord = new javax.swing.JButton();
        comboPessoas = new javax.swing.JComboBox<>();
        btnEditar = new javax.swing.JButton();
        btnExcluir = new javax.swing.JButton();
        btnSalvar = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Banco de Dados SQL/Excel/Word");

        jPanel1.setToolTipText("Banco de dados -SQL - Word- Excel");

        btnExcel.setText("Excel");
        btnExcel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnExcelActionPerformed(evt);
            }
        });

        btnInserir.setBackground(new java.awt.Color(102, 255, 102));
        btnInserir.setText("Inserir");
        btnInserir.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnInserirActionPerformed(evt);
            }
        });

        lblNome.setText("Nome:");

        lblIdade.setText("Idade:");

        lblProfissao.setText("Profissão:");

        btnWord.setText("World");
        btnWord.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnWordActionPerformed(evt);
            }
        });

        comboPessoas.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));
        comboPessoas.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                comboPessoasActionPerformed(evt);
            }
        });

        btnEditar.setBackground(new java.awt.Color(255, 255, 51));
        btnEditar.setText("Editar");
        btnEditar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnEditarActionPerformed(evt);
            }
        });

        btnExcluir.setBackground(new java.awt.Color(255, 51, 51));
        btnExcluir.setText("Excluir");
        btnExcluir.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnExcluirActionPerformed(evt);
            }
        });

        btnSalvar.setBackground(new java.awt.Color(102, 255, 102));
        btnSalvar.setText("Salvar");
        btnSalvar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnSalvarActionPerformed(evt);
            }
        });

        jLabel1.setText("Lista de nomes Salvas em SQL:");

        jLabel2.setText("Exportar para:");

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(comboPessoas, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addContainerGap()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel1)
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addComponent(jLabel2)
                                        .addGap(18, 18, 18)
                                        .addComponent(btnExcel)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                        .addComponent(btnWord)))
                                .addGap(0, 0, Short.MAX_VALUE))
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                .addGap(0, 24, Short.MAX_VALUE)
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                        .addComponent(lblNome)
                                        .addGap(18, 18, 18)
                                        .addComponent(txtNome, javax.swing.GroupLayout.PREFERRED_SIZE, 294, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                        .addComponent(lblIdade)
                                        .addGap(18, 18, 18)
                                        .addComponent(txtIdade, javax.swing.GroupLayout.PREFERRED_SIZE, 294, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                        .addComponent(lblProfissao)
                                        .addGap(18, 18, 18)
                                        .addComponent(txtProfissao, javax.swing.GroupLayout.PREFERRED_SIZE, 294, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                        .addComponent(btnInserir)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(btnSalvar)
                                        .addGap(18, 18, 18)
                                        .addComponent(btnEditar)
                                        .addGap(18, 18, 18)
                                        .addComponent(btnExcluir)))))))
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addGap(20, 20, 20)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lblNome)
                    .addComponent(txtNome, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lblIdade)
                    .addComponent(txtIdade, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lblProfissao)
                    .addComponent(txtProfissao, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(btnExcel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btnWord, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jLabel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnInserir)
                    .addComponent(btnEditar)
                    .addComponent(btnExcluir)
                    .addComponent(btnSalvar))
                .addGap(18, 18, 18)
                .addComponent(jLabel1)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(comboPessoas, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(68, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 13, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );

        setBounds(0, 0, 428, 301);
    }// </editor-fold>//GEN-END:initComponents

    private void btnInserirActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnInserirActionPerformed
        try {
            String nome = txtNome.getText();
            int idade = Integer.parseInt(txtIdade.getText());
            String profissao = txtProfissao.getText();

            pessoaDAO.inserirPessoa(nome, idade, profissao);
            atualizarComboPessoas();

            txtNome.setText("");
            txtIdade.setText("");
            txtProfissao.setText("");

            JOptionPane.showMessageDialog(null, "Pessoa inserida com sucesso!");
        } catch (NumberFormatException ex) {
            JOptionPane.showMessageDialog(null, "Por favor, insira uma idade válida.");
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Erro ao inserir pessoa: " + ex.getMessage());
        }


    }//GEN-LAST:event_btnInserirActionPerformed

    private void btnExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnExcelActionPerformed
        try {
            PessoaDAO dao = new PessoaDAO();
            List<Pessoa> pessoas = dao.listarPessoas();

            ExportarExcel exportar = new ExportarExcel();

            // Caminho para o arquivo original
            String caminhoArquivo = "Arquivos modelos Word e Excel/pessoas.xlsx"; // Caminho relativo
            exportar.exportarParaExcel(pessoas, caminhoArquivo);

            // Formata a data e a hora atual para o nome do arquivo backup
            SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy HH-mm-ss");
            String dataAtual = sdf.format(new java.util.Date());

            // Caminho para a pasta de backup
            String caminhoBackupPasta = "Arquivos Excel Backup";
            Files.createDirectories(Paths.get(caminhoBackupPasta)); // Cria a pasta se não existir

            // Caminho para o arquivo backup
            String caminhoBackupArquivo = caminhoBackupPasta + "/pessoas_" + dataAtual + ".xlsx";
            Files.copy(Paths.get(caminhoArquivo), Paths.get(caminhoBackupArquivo), StandardCopyOption.REPLACE_EXISTING);

            // Mensagem de sucesso
            javax.swing.JOptionPane.showMessageDialog(this, "Dados exportados para Excel com sucesso! O arquivo está localizado em: " + caminhoArquivo + "\nCópia de backup em: " + caminhoBackupArquivo);

        } catch (Exception e) {
            javax.swing.JOptionPane.showMessageDialog(this, "Erro ao exportar para Excel: " + e.getMessage());
        }
        /* try {
            PessoaDAO dao = new PessoaDAO();
            List<Pessoa> pessoas = dao.listarPessoas();

            ExportarExcel exportar = new ExportarExcel();

            // Altere o caminho para relativo
            String caminhoArquivo = "Arquivos modelos Word e Excel/pessoas.xlsx"; // Caminho relativo
            exportar.exportarParaExcel(pessoas, caminhoArquivo);

            javax.swing.JOptionPane.showMessageDialog(this, "Dados exportados para Excel com sucesso! O arquivo está localizado em: " + caminhoArquivo);

        } catch (Exception e) {
            javax.swing.JOptionPane.showMessageDialog(this, "Erro ao exportar para Excel: " + e.getMessage());
        }
         */
    }//GEN-LAST:event_btnExcelActionPerformed

    private void btnWordActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnWordActionPerformed
        try {
            // Obtém o nome da pessoa selecionada no combo box
            String selecionado = (String) comboPessoas.getSelectedItem();

            // Verifica se uma pessoa foi selecionada
            if (selecionado != null) {
                // Divide a string para obter nome, idade e profissão
                String[] partes = selecionado.split(" - ");
                String nome = partes[0]; // Nome
                int idade = Integer.parseInt(partes[1]); // Idade
                String profissao = partes[2]; // Profissão

                // Caminho do arquivo modelo
                String caminhoModelo = "Arquivos modelos Word e Excel/pessoas.docx";// Caminho relativo do modelo
                // Formata a data atual
                SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd_HHmmss");
                String dataAtual = sdf.format(new java.util.Date());

                // Cria um nome de arquivo com o nome da pessoa e a data
                String novoArquivo = "Arquivos Word gerados/" + nome + "_" + dataAtual + ".docx";
                Files.copy(Paths.get(caminhoModelo), Paths.get(novoArquivo), StandardCopyOption.REPLACE_EXISTING);

                // Abre o novo arquivo para preencher
                try (XWPFDocument documento = new XWPFDocument(new FileInputStream(novoArquivo))) {

                    // Substitui os marcadores no documento
                    for (XWPFParagraph paragrafo : documento.getParagraphs()) {
                        for (XWPFRun run : paragrafo.getRuns()) {
                            String texto = run.getText(0);
                            if (texto != null) {
                                // Substituições dos marcadores
                                texto = texto.replace("{{nome}}", nome);
                                texto = texto.replace("{{idade}}", String.valueOf(idade));
                                texto = texto.replace("{{profissao}}", profissao);
                                run.setText(texto, 0); // Atualiza o texto do run
                            }
                        }
                    }

                    // Salva o documento preenchido
                    try (FileOutputStream out = new FileOutputStream(novoArquivo)) {
                        documento.write(out);
                    }
                }

                // Mensagem de sucesso
                javax.swing.JOptionPane.showMessageDialog(this, "Dados completados e exportados para Word com sucesso! O arquivo está localizado em: " + novoArquivo);
            } else {
                javax.swing.JOptionPane.showMessageDialog(this, "Nenhuma pessoa selecionada.");
            }

        } catch (Exception e) {
            javax.swing.JOptionPane.showMessageDialog(this, "Erro ao completar e exportar para Word: " + e.getMessage());
        }
    }//GEN-LAST:event_btnWordActionPerformed

    private void comboPessoasActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_comboPessoasActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_comboPessoasActionPerformed

    private void btnSalvarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnSalvarActionPerformed

        String nome = txtNome.getText();
        int idade = Integer.parseInt(txtIdade.getText());
        String profissao = txtProfissao.getText();
        String selecionado = (String) comboPessoas.getSelectedItem();

        if (selecionado != null) {
            try {
                String[] partes = selecionado.split(" - ");
                String nomeOriginal = partes[0];
                int idadeOriginal = Integer.parseInt(partes[1]);
                String profissaoOriginal = partes[2];

                // Buscar o ID baseado nos dados originais
                int id = pessoaDAO.buscarIdPorDados(nomeOriginal, idadeOriginal, profissaoOriginal);

                if (id > 0) {
                    pessoaDAO.editarPessoa(id, nome, idade, profissao);
                    // Atualiza a interface após salvar
                    atualizarComboPessoas();
                    txtNome.setText("");
                    txtIdade.setText("");
                    txtProfissao.setText("");
                    btnSalvar.setVisible(false); // Oculta o botão "Salvar"
                    btnInserir.setVisible(true); // Exibe o botão "Inserir"
                    JOptionPane.showMessageDialog(null, "Pessoa atualizada com sucesso!");
                } else {
                    JOptionPane.showMessageDialog(null, "Não foi possível encontrar o ID da pessoa selecionada.");
                }
            } catch (NumberFormatException e) {
                JOptionPane.showMessageDialog(null, "Erro ao converter valores: " + e.getMessage());
            } catch (IllegalArgumentException e) {
                JOptionPane.showMessageDialog(null, "Erro: " + e.getMessage());
            }
        }
    }//GEN-LAST:event_btnSalvarActionPerformed

    private void btnEditarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnEditarActionPerformed
        String selecionado = (String) comboPessoas.getSelectedItem();

        if (selecionado != null) {
            String[] partes = selecionado.split(" - ");
            String nome = partes[0];
            int idade = Integer.parseInt(partes[1]);
            String profissao = partes[2];

            txtNome.setText(nome);
            txtIdade.setText(String.valueOf(idade));
            txtProfissao.setText(profissao);

            btnInserir.setVisible(false);  // Oculta o botão "Inserir"
            btnSalvar.setVisible(true);    // Exibe o botão "Salvar"
        } else {
            JOptionPane.showMessageDialog(this, "Nenhuma pessoa selecionada.");
        }

    }//GEN-LAST:event_btnEditarActionPerformed

    private void btnExcluirActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnExcluirActionPerformed
        String selecionado = (String) comboPessoas.getSelectedItem();
        if (selecionado != null) {
            try {
                String[] partes = selecionado.split(" - ");
                // Obter os dados da pessoa selecionada
                String nome = partes[0];
                int idade = Integer.parseInt(partes[1]);
                String profissao = partes[2];

                // Confirmar antes de excluir
                int confirma = JOptionPane.showConfirmDialog(
                        this,
                        "Tem certeza que deseja excluir " + nome + "?",
                        "Confirmar Exclusão",
                        JOptionPane.YES_NO_OPTION
                );

                if (confirma == JOptionPane.YES_OPTION) {
                    // Buscar o ID baseado nos dados originais
                    int id = pessoaDAO.buscarIdPorDados(nome, idade, profissao);

                    if (id > 0) {
                        pessoaDAO.excluirPessoa(id);
                        atualizarComboPessoas();
                        txtNome.setText("");
                        txtIdade.setText("");
                        txtProfissao.setText("");
                        JOptionPane.showMessageDialog(null, "Pessoa excluída com sucesso!");
                    } else {
                        JOptionPane.showMessageDialog(null, "Não foi possível encontrar o ID da pessoa selecionada.");
                    }
                }
            } catch (NumberFormatException e) {
                JOptionPane.showMessageDialog(null, "Erro ao converter valores: " + e.getMessage());
            } catch (Exception e) {
                JOptionPane.showMessageDialog(null, "Erro ao excluir: " + e.getMessage());
            }
        } else {
            JOptionPane.showMessageDialog(this, "Nenhuma pessoa selecionada.");
        }
    }//GEN-LAST:event_btnExcluirActionPerformed

    private void atualizarComboPessoas() {
        try {
            // Obter a lista de pessoas do banco de dados
            List<Pessoa> pessoas = pessoaDAO.listarPessoas();
            comboPessoas.removeAllItems(); // Limpar o JComboBox

            // Adicionar os itens no formato "nome - idade - profissão"
            for (Pessoa pessoa : pessoas) {
                String itemCombo = pessoa.getNome() + " - " + pessoa.getIdade() + " - " + pessoa.getProfissao();
                comboPessoas.addItem(itemCombo);
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Erro ao carregar pessoas: " + ex.getMessage());
        }
    }

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
            java.util.logging.Logger.getLogger(cadastroApp.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(cadastroApp.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(cadastroApp.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(cadastroApp.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new cadastroApp().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnEditar;
    private javax.swing.JButton btnExcel;
    private javax.swing.JButton btnExcluir;
    private javax.swing.JButton btnInserir;
    private javax.swing.JButton btnSalvar;
    private javax.swing.JButton btnWord;
    private javax.swing.JComboBox<String> comboPessoas;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JLabel lblIdade;
    private javax.swing.JLabel lblNome;
    private javax.swing.JLabel lblProfissao;
    private javax.swing.JTextField txtIdade;
    private javax.swing.JTextField txtNome;
    private javax.swing.JTextField txtProfissao;
    // End of variables declaration//GEN-END:variables
}
