# Exemplo de Suplemento do Painel de Tarefas do gerador de CSV para o Excel 2016

_Aplica-se ao: Excel 2016_

Esse suplemento do painel de tarefas mostra como criar uma tabela a partir de uma lista de nomes da coluna usando as APIs JavaScript no Excel 2016. Há dois tipos: o editor de código e o Visual Studio.

![Exemplo de Gerador de CSV](../Images/ScreenCap1.PNG)

## Experimente
### Versão do editor de código

A maneira mais fácil de implantar e testar o suplemento é copiar os arquivos em um compartilhamento de rede.

1.  Crie uma pasta em um compartilhamento de rede (por exemplo, \\\MyShare\Excel_CSV_Generator) e copie todos os arquivos para a pasta do Editor de Código. 
2.  Edite o elemento <SourceLocation> do arquivo de manifesto de modo que ele aponte para o local de compartilhamento criado na etapa 1. 
3.  Copie o manifesto (TeacherCSVGenerator.xml) para um compartilhamento de rede (por exemplo, \\\MyShare\\MyManifests).
4.  Adicione o local de compartilhamento que contém o manifesto como um catálogo de aplicativos confiáveis no Excel.

    a. Inicie o Excel e abra uma planilha em branco.  
    
    b. Escolha a guia **Arquivo** e escolha **Opções**.
    
    c. Escolha **Central de Confiabilidade** e, em seguida, escolha o botão **Configurações da Central de Confiabilidade**.
    
    d. Escolha **Catálogos de Suplementos Confiáveis**.
    
    e. Na caixa  **URL de Catálogo**, insira o caminho para o compartilhamento de rede que você criou na etapa 3 e escolha **Adicionar Catálogo**.
    
   f.  Marque a caixa de seleção **Mostrar no Menu** e escolha **OK**. Será exibida uma mensagem para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Office. 
        
5.  Teste e execute o suplemento. 

    a. Na **guia Inserir** no Excel 2016, escolha **Meus Suplementos**. 
    
    b. Na caixa de diálogo **Suplementos do Office**, escolha **Pasta Compartilhada**.
    
    c. Escolha **Amostra de Lista de Participação em Aula para Professores no formato CVS**>**Inserir**. O suplemento abre um painel de tarefas e cria a lista de participaçã	o em aula no formato CSV na planilha ativa, conforme mostrado nesta captura de tela. 
      
   ![Exemplo de Controlador de Orçamento Escolar](../Images/ScreenCap2.PNG) 

    d. Escolha um serviço de gerenciamento de sala de aula.
    
    e. Clique no botão Fazer Lista de Participação para inserir uma lista vazia na planilha ativa  
    
      ![Amostra de Controlador de Orçamento Escolar](../Images/ScreenCap3.PNG) 
      
    f. Clique no botão Ajuda de Exportação do Excel para saber como exportar uma planilha como um arquivo .csv.  
  
    
### Versão do Visual Studio
1.  Copie o projeto para uma pasta local e abra o TeacherCSVGenerator.sln no Visual Studio.
2.  Pressione F5 para criar e implantar o suplemento de exemplo. O Excel inicia e o suplemento abre em um painel de tarefas à direita da planilha em branco, conforme mostrado na captura de tela a seguir. 
        
  ![Exemplo de Gerador de CSV do Excel](../Images/ScreenCap1.PNG) 

3.  Escolha um serviço de gerenciamento da sala de aula online na lista suspensa
4.  Adicione uma tabela de lista de participação dos alunos usando o botão **Criar Lista de Participação** e veja a tabela criada na planilha ativa.

  ![Exemplo de Controlador de Orçamento Escolar](../Images/ScreenCap3.PNG) 
5.  Adicione alunos à lista de participação preenchendo as células nas linhas abaixo do cabeçalho da tabela.
6.  Use o recurso de exportação no Excel para salvar a planilha como um arquivo .csv. Esse arquivo está no formato correto para ser importado para o serviço de sua escolha.


### Saiba mais

As APIs JavaScript para Excel têm muito mais a oferecer à medida que você desenvolve suplementos. Confira a seguir alguns dos recursos disponíveis. 

1.  [Visão geral da programação de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorador de trecho para Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Exemplos de código de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [Referência da API JavaScript para Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Crie seu primeiro Suplemento do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)

