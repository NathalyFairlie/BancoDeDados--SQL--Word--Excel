package NathalyFairlie.wordexcel1.office;

import NathalyFairlie.wordexcel1.model.Pessoa;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.util.List;

public class ExportarWord {
    public void exportarParaWord(List<Pessoa> pessoas, String caminhoArquivo) throws Exception {
        XWPFDocument document = new XWPFDocument();

        for (Pessoa pessoa : pessoas) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText("- " + pessoa.getNome() + ", " + pessoa.getIdade() + " anos");
        }

        try (FileOutputStream fileOut = new FileOutputStream(caminhoArquivo)) {
            document.write(fileOut);
        }

        document.close();
    }
}
