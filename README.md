# BancoDeDados--SQL--Word--Excel

## Descrição

Programa Java/SQL para armazenar e gerenciar informações de Nome, Idade e Profissão de pessoas em um banco de dados SQL. Permite a geração do banco de dados em Excel e o preenchimento automático de formulários Word.

## Funcionalidades Gerais

*   **CRUD:** Inserir, editar, excluir e visualizar registros de pessoas (nome, idade, profissão).
*   **Exportação para Excel:** Exporta os dados da tabela SQL para um arquivo Excel (`.xlsx`).
*   **Geração de Documento Word:** Preenche um modelo de documento Word (`.docx`) com os dados de uma pessoa selecionada, substituindo marcadores (como `{{nome}}`, `{{idade}}`, `{{profissao}}`).
*   **Interface Gráfica:** Interface intuitiva para interagir com o banco de dados e as funcionalidades de exportação.

## Funcionalidades Principais

*   **Inserir:** Adiciona uma nova pessoa ao banco de dados.
*   **Editar:** Modifica os dados de uma pessoa existente.
*   **Excluir:** Remove uma pessoa do banco de dados.
*   **Excel:** Exporta os dados para um arquivo Excel. (As alterações no BD afetam diretamente o arquivo Excel quando o botão é clicado).
*   **Word:** Gera um novo documento Word preenchido com os dados da pessoa selecionada a partir de um modelo Word.
*   **Combo Box:** Exibe uma lista de pessoas cadastradas no banco de dados.

## Uso

Insira dados (nome, idade, profissão) e salve-os no banco de dados SQL. Visualize e selecione dados através do `JComboBox`. Edite, exclua ou exporte para Excel/Word.

## Notas

*   Os caminhos dos arquivos Excel e Word são relativos. As pastas "Arquivos modelos Word e Excel" e "Arquivos Word gerados" devem estar na mesma pasta do executável.
