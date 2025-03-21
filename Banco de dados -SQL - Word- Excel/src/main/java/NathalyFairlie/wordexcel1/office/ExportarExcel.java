package bNathalyFairlie.wordexcel1.office;

import NathalyFairlie.wordexcel1.model.Pessoa;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class ExportarExcel {

    public void exportarParaExcel(List<Pessoa> pessoas, String caminhoArquivo) throws Exception {
        // Carregar o arquivo Excel existente
        FileInputStream fis = new FileInputStream(caminhoArquivo);
        Workbook workbook = new XSSFWorkbook(fis);

        // Localizar a aba do banco de dados (ou criar uma nova se não existir)
        Sheet sheetBancoDados = workbook.getSheet("Banco de Dados");
        if (sheetBancoDados == null) {
            sheetBancoDados = workbook.createSheet("Banco de Dados");
        }

        // Limpar a aba do banco de dados antes de preencher
        int numberOfRows = sheetBancoDados.getPhysicalNumberOfRows();
        for (int i = 0; i < numberOfRows; i++) {
            Row row = sheetBancoDados.getRow(i);
            if (row != null) {
                sheetBancoDados.removeRow(row);
            }
        }

        // Cabeçalhos para a aba "Banco de Dados"
        Row headerRow = sheetBancoDados.createRow(0);
        headerRow.createCell(0).setCellValue("Nome");
        headerRow.createCell(1).setCellValue("Idade");
        headerRow.createCell(2).setCellValue("Profissão");

        // Adiciona as pessoas na aba "Banco de Dados"
        int rowNum = 1;
        for (Pessoa pessoa : pessoas) {
            Row row = sheetBancoDados.createRow(rowNum++);
            row.createCell(0).setCellValue(pessoa.getNome());
            row.createCell(1).setCellValue(pessoa.getIdade());
            row.createCell(2).setCellValue(pessoa.getProfissao());
        }

        // Preencher a aba "Recibo" com os dados da primeira pessoa, por exemplo
        Sheet sheetRecibo = workbook.getSheet("Recibo");
        if (sheetRecibo != null && !pessoas.isEmpty()) {
            Pessoa pessoa = pessoas.get(0);
            Row row = sheetRecibo.getRow(1); // Supondo que a linha 1 seja onde o recibo é preenchido
            if (row == null) {
                row = sheetRecibo.createRow(1);
            }
            row.createCell(0).setCellValue("Olá, meu nome é " + pessoa.getNome());
            row.createCell(1).setCellValue("Tenho " + pessoa.getIdade() + " anos");
            row.createCell(2).setCellValue("Atualmente trabalho como " + pessoa.getProfissao());
        }

        // Fechar o FileInputStream
        fis.close();

        // Salvar o arquivo Excel modificado
        try (FileOutputStream fos = new FileOutputStream(caminhoArquivo)) {
            workbook.write(fos);
        }
        workbook.close();
    }
}
