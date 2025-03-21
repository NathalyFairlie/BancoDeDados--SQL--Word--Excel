
package NathalyFairlie.wordexcel1.model;

public class Pessoa {
    private int id;
    private String nome;
    private int idade;
    private String profissao;

    public Pessoa(int id, String nome, int idade, String profissao) {
        this.id = id;
        this.nome = nome;
        this.idade = idade;
        this.profissao = profissao;
    }

    // Getters e setters

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getNome() {
        return nome;
    }

    public void setNome(String nome) {
        this.nome = nome;
    }

    public int getIdade() {
        return idade;
    }

    public void setIdade(int idade) {
        this.idade = idade;
    }

    public String getProfissao() {
        return profissao;
    }

    public void setProfissao(String profissao) {
        this.profissao = profissao;
    }
    
}

