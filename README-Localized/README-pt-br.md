---
page_type: sample
products:
- office-excel
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 10/15/2015 1:50:50 PM
---
# <a name="csv-generator-task-pane-add-in-sample-for-excel-2016"></a>Exemplo do Suplemento do Painel de Tarefas do gerador de CSV para o Excel 2016

_Aplica-se a: Excel 2016_

Esse suplemento do painel de tarefas mostra como criar uma tabela a partir de uma lista de nomes da coluna usando as APIs JavaScript no Excel 2016. Há dois tipos: o editor de código e o Visual Studio.

![Exemplo de Gerador de CSV](../Images/ScreenCap1.PNG)

## <a name="try-it-out"></a>Experimente
### <a name="code-editor-version"></a>Versão do editor de código

A maneira mais fácil de implantar e testar o suplemento é copiar os arquivos para um compartilhamento de rede.

1.  Hospede os arquivos na pasta projeto do Editor de Códigos usando um servidor de sua escolha.
2.  Edite os elementos \<SourceLocation\> e \<URL\> do arquivo de manifesto para que ele aponte para o local hospedado criado na etapa 1. (por exemplo, https://localhost/CSVGenerator/Home.html)
3.  Copie o manifesto (TeacherCSVGenerator.xml) para um compartilhamento de rede (por exemplo, \\\MyShare\\MyManifests).
4.  Adicione o local de compartilhamento que contém o manifesto como um catálogo de aplicativos confiáveis no Excel.

    a.  Inicie o Excel e abra uma planilha em branco.

    b.  Escolha a guia **Arquivo** e escolha **Opções**.

    c.  Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.

    d.  Escolha **Catálogos de Suplemento Confiáveis**.

    e.  Na caixa **URL de Catálogo**, insira o caminho para o compartilhamento de rede que você criou na etapa 3 e escolha **Adicionar Catálogo**.

   f.  Marque a caixa de seleção **Mostrar no Menu** e escolha **OK**. Será exibida uma mensagem para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Office.

5.  Teste e execute o suplemento.

    a.  Na **guia Inserir** no Excel 2016, escolha **Meus Suplementos**.

    b.  Na caixa de diálogo **Suplementos do Office**, escolha **Pasta Compartilhada**.

    c.  Escolha **Exemplo de Lista de Participação em Aula para Professores no formato CSV**>**Inserir**. O suplemento abre um painel de tarefas e cria a lista de participação em aula no formato CSV na planilha ativa, conforme mostrado nesta captura de tela.

   ![Amostra de Controlador de Orçamento Escolar](../Images/ScreenCap2.PNG)

    d.  Escolha um serviço de gerenciamento de sala de aula.

    e.  Clique no botão Fazer Lista de Participação para inserir uma lista vazia na planilha ativa

      ![Amostra de Controlador de Orçamento Escolar](../Images/ScreenCap3.PNG)

    f.  Clique no botão Ajuda de Exportação do Excel para saber como exportar uma planilha como um arquivo .csv.


### <a name="visual-studio-version"></a>Versão do Visual Studio
1.  Copie o projeto para uma pasta local e abra o TeacherCSVGenerator.sln no Visual Studio.
2.  Pressione F5 para criar e implantar o suplemento de exemplo. O Excel inicia e o suplemento abre em um painel de tarefas à direita da planilha em branco, conforme mostrado na captura de tela a seguir.

  ![Exemplo de Gerador de CSV do Excel](../Images/ScreenCap1.PNG)

3.  Escolha um serviço de gerenciamento da sala de aula online na lista suspensa
4.  Adicione uma tabela de lista de participação dos alunos usando o botão **Criar Lista de Participação** e veja a tabela criada na planilha ativa.

  ![Amostra de Controlador de Orçamento Escolar](../Images/ScreenCap3.PNG)
5.  Adicione alunos à lista de participação preenchendo as células nas linhas abaixo do cabeçalho da tabela.
6.  Use o recurso de exportação no Excel para salvar a planilha como um arquivo .csv. Esse arquivo está no formato correto para ser importado para o serviço de sua escolha.


### <a name="learn-more"></a>Saiba mais

As APIs JavaScript para Excel têm muito mais a oferecer à medida que você desenvolve suplementos. Confira a seguir alguns dos recursos disponíveis.

1.  [Visão geral da programação de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorador de Trechos para Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Exemplos de código de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [Referência da API JavaScript de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Criar seu primeiro Suplemento do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)


Este projeto adotou o [Código de Conduta de Software Livre da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira as [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.
