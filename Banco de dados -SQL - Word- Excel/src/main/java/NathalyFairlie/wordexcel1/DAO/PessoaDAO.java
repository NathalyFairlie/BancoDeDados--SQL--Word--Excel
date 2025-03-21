package NathalyFairlie.wordexcel1.DAO;

import NathalyFairlie.wordexcel1.model.Pessoa;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

public class PessoaDAO {

    private final String url = "jdbc:mysql://localhost:3306/cadastrodb"; // Nome correto do banco de dados
    private final String usuario = "root"; // Usuário do MySQL
    private final String senha = ""; // Senha do MySQL

    // Método para buscar todas as pessoas no banco de dados
    public List<Pessoa> listarPessoas() {
        List<Pessoa> pessoas = new ArrayList<>();
        String sql = "SELECT * FROM pessoas"; // Ajuste na tabela

        try (Connection conn = DriverManager.getConnection(url, usuario, senha); PreparedStatement stmt = conn.prepareStatement(sql); ResultSet rs = stmt.executeQuery()) {

            while (rs.next()) {
                int id = rs.getInt("id");
                String nome = rs.getString("nome");
                int idade = rs.getInt("idade");
                String profissao = rs.getString("profissao");
                pessoas.add(new Pessoa(id, nome, idade, profissao));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }

        return pessoas;
    }

    // Método para inserir uma pessoa no banco de dados
    public void inserirPessoa(String nome, int idade, String profissao) {
        // Validação e formatação dos campos antes de salvar no banco de dados
        nome = formatarNomeProfissao(nome);
        profissao = formatarNomeProfissao(profissao);

        if (!validarNomeProfissao(nome) || !validarNomeProfissao(profissao)) {
            throw new IllegalArgumentException("Nome ou profissão inválidos. Somente letras e espaços únicos são permitidos.");
        }

        if (!validarIdade(idade)) {
            throw new IllegalArgumentException("Idade inválida. Deve estar entre 1 e 99.");
        }

        String sql = "INSERT INTO pessoas (nome, idade, profissao) VALUES (?, ?, ?)";
        try (Connection conn = DriverManager.getConnection(url, usuario, senha); PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, nome);
            stmt.setInt(2, idade);
            stmt.setString(3, profissao);
            stmt.executeUpdate();
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    // Método para formatar "nome" e "profissao" com a primeira letra maiúscula e espaçamento correto
    private String formatarNomeProfissao(String texto) {
        String[] palavras = texto.trim().split("\\s+");
        StringBuilder resultado = new StringBuilder();

        for (String palavra : palavras) {
            resultado.append(Character.toUpperCase(palavra.charAt(0)))
                    .append(palavra.substring(1).toLowerCase())
                    .append(" ");
        }

        return resultado.toString().trim();
    }

    // Valida se o "nome" e "profissao" contém apenas letras e espaços
    private boolean validarNomeProfissao(String texto) {
        return texto.matches("^[A-Za-zÀ-ÖØ-öø-ÿ]+(\\s[A-Za-zÀ-ÖØ-öø-ÿ]+)*$");
    }

    // Valida se a "idade" está entre 1 e 99 e contém apenas números
    private boolean validarIdade(int idade) {
        return idade >= 1 && idade <= 99;
    }

    // Método para editar uma pessoa no banco de dados
    public void editarPessoa(int id, String nome, int idade, String profissao) {
        // Validação e formatação dos campos antes de atualizar no banco de dados
        nome = formatarNomeProfissao(nome);
        profissao = formatarNomeProfissao(profissao);

        if (!validarNomeProfissao(nome) || !validarNomeProfissao(profissao)) {
            throw new IllegalArgumentException("Nome ou profissão inválidos. Somente letras e espaços únicos são permitidos.");
        }

        if (!validarIdade(idade)) {
            throw new IllegalArgumentException("Idade inválida. Deve estar entre 1 e 99.");
        }

        String sql = "UPDATE pessoas SET nome = ?, idade = ?, profissao = ? WHERE id = ?";
        try (Connection conn = DriverManager.getConnection(url, usuario, senha); PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, nome);
            stmt.setInt(2, idade);
            stmt.setString(3, profissao);
            stmt.setInt(4, id);
            stmt.executeUpdate();
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    // Método para excluir uma pessoa do banco de dados
    public void excluirPessoa(int id) {
        String sql = "DELETE FROM pessoas WHERE id = ?";
        try (Connection conn = DriverManager.getConnection(url, usuario, senha); PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setInt(1, id);
            stmt.executeUpdate();
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    public int buscarIdPorDados(String nome, int idade, String profissao) {
        String sql = "SELECT id FROM pessoas WHERE nome = ? AND idade = ? AND profissao = ? LIMIT 1";
        try (Connection conn = DriverManager.getConnection(url, usuario, senha); PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, nome);
            stmt.setInt(2, idade);
            stmt.setString(3, profissao);

            try (ResultSet rs = stmt.executeQuery()) {
                if (rs.next()) {
                    return rs.getInt("id");
                }
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return -1; // ID não encontrado
    }
}
