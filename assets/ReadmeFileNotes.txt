# excel-to-xml-converter

Allow to convert Excel `.xlsx` files into `.xml` files | This example was built to simulate report issuing to the Angola Central Bank

## 🚀 Beginning

The purpose of this project is to create a script that can read an Excel file `XLSX` and displays its contents in `XML` format.

### 🔧 Prerequisites

To develop the project, the following tools must first be installed:

- Code editor: **[Visual Studio Code](https://code.visualstudio.com/)** ![VS Code](/assets/img/vs_code_logo.png "VS Code")
- `Windows` `GCC` Compiler: **[MinGW-w64](https://www.mingw-w64.org/downloads/)** ![MinGW-w64](/assets/img/gcc_compiler.png "MinGW-w64 GCC Compiler")
- Support library (`XLSX` reading): **[libxlsxio](https://sourceforge.net/projects/xlsxio/files/0.2.31/xlsxio-0.2.31-binary-win64.zip/download/)**
- Compiler for support library **`libxlsio`** (optional): **[CMake](https://cmake.org/download/)** ![CMake](/assets/img/cmake.png "CMake")
- VS Code extension to support IntelliSense and build, etc: **[Microsoft C/C++ Extension for VS Code](https://marketplace.visualstudio.com/items?itemName=ms-vscode.cpptools)** ![C/C++ VS Code Extension](/assets/img/vs_code_extension.png "VS Code Extension")
- **[Git](https://git-scm.com/downloads)** ![Git](/assets/img/git_logo.png "Git")

❗ **Important**: After installing MinGW, add the bin path to the system PATH.

```makefile
C:\mingw-w64\bin
```

### ⌨️ Code

Executar o SQL Server Management Studio

![ssms1](/assets/img/ssms1.png "SSMS")

Efectuar a conexão à instância

![ssms2](/assets/img/ssms2.png "SSMS")

Escolher a base de dados de trabalho

![ssms3](/assets/img/ssms3.png "SSMS")

Montar a *query*

![ssms4](/assets/img/ssms4.png "SSMS")

Após abrir o Visual Studio, ao criar o novo projecto, pesquisar pela extensão "Primavera" e escolher a opção desejada.

![vs-project](/CegidPrimaveraExtensibilitySamp1/Base/Assets/img/vs_project.png "Project")

Após escolher o tipo de linguagem a utilizar, clicar em "*Next*" ou "Próximo". Na página seguinte, escolher o nome do projecto, a versão do pacote .NET Framework a utilizar e clicar em "*Create*" ou "Criar".
Na janela que aparecer, devemos escolher sobre qual componente Primavera iremos trabalhar.

**Nota**: para que esta última janela apareça, é necessário que a configuração do caminho de **[instalação](#-instala%C3%A7%C3%A3o)** "*Installation Path*" esteja bem feita.

Para este projecto, clicou-se nas opções de edição da "Ficha de Artigos" e utilização dos motores (modelo e controlador) para artigos, como se vê nas imagens abaixo.

![project-option1](/CegidPrimaveraExtensibilitySamp1/Base/Assets/img/project_option1.png "Project Options")
![project-option2](/CegidPrimaveraExtensibilitySamp1/Base/Assets/img/project_option2.png "Project Options")
![project-option3](/CegidPrimaveraExtensibilitySamp1/Base/Assets/img/project_option3.png "Project Options")

#### Diagrama de Arquitectura

A seguir, é apresentado o diagrama de arquitectura do projecto (escrito com **[Mermaid](https://mermaid.js.org/)**), destacando a separação das responsabilidades entre as camadas. Desde a interface do utilizador até aos mecanismos de interação com sistemas externos, cada elemento é estrategicamente posicionado para reforçar a modularidade, a escalabilidade e a manutenibilidade do sistema. Esta estrutura facilita a compreensão de como os componentes colaboram para a realização dos objectivos do software, alinhando-se assim, aos princípios de "***Clean Code***" e das boas práticas de programação em projectos de colaboração.

##### Diagrama 1

![arch-diagram](/CegidPrimaveraExtensibilitySamp1/Base/Assets/img/arch_diagram.png "Architecture Diagram")

##### Diagrama com Mermaid

```mermaid
graph LR;
    subgraph layer-infra[Infraestrutura];
        tests("Testes (Tests)")
        exceptions("Excepções (Exceptions)")
        models("Modelo (Models)")
            models --> bd("Base de Dados") <--> exceptions
        controllers("Controlador (Controllers)") --> models
            controllers --> business("Lógica de Negócios") <--> exceptions
        views("Apresentação (Views)") --> controllers
    end

classDef infra fill:#D4E6F1, stroke:#2980B9, color:#154360, stroke-width:2px; 
classDef view fill:#D5F5E3, stroke:#27AE60, color:#145A32, stroke-width:2px;      %% Verde para view
classDef control fill:#FADBD8, stroke:#E74C3C, color:#78281F, stroke-width:2px;   %% Rosa para controller
classDef model fill:#EBDEF0, stroke:#8E44AD, color:#4A235A, stroke-width:2px;     %% Roxo para model
classDef exception fill:#FDEBD0, stroke:#F39C12, color:#7E5109, stroke-width:2px; %% Amarelo para exception
classDef test fill:#8BD3DD, stroke:#006B7D, color:#006B7D, stroke-width:2px;      %% Azul para test

class infra infra;
class views view;
class business,controllers control;
class bd,models model;
class exceptions exception;
class tests test
```

Os diagramas acima espelhados mostram as seguintes interacções:

- Que o utilizador interage com a **Apresentação** (Views);
- A Apresentação envia requisições ao **Controlador** (Controllers);
- O Controlador invoca os métodos do **Modelo** (Models) para realizar as operações necessárias;
- O Modelo contém a classe de acesso à **Base de Dados**;
- A **Lógica de Negócios** é utilizada pelo Controlador para aplicar regras específicas;
- As **Excepções** (Exceptions) são lançadas pelas classes de Lógica de Negócios e Base de Dados para sinalizar erros e são tratados em uma camada superior; 
- Os **Testes** (Tests) interagem com todos os componentes para verificar seu bom funcionamento.

**Obs**: Os diagramas são uma representação simplificada da arquitectura deste projecto, a complexidade do diagrama pode variar dependendo do tamanho e da complexidade do projecto. Outros componentes e interacções podem ser adicionados conforme a necessidade; eles proporcionam uma visão geral dos principais componentes e suas interacções e podem ser úteis para que se entenda a estrutura do sistema, comunicar a arquitectura para outras pessoas e planear o desenvolvimento de futuras funcionalidades (***features***).

#### Estrutura de Pastas

Em acordo com a organização apresentada no diagrama de arquitectura, a estrutura de pastas do projecto sugere uma arquitectura **MVC** (*Models*, *Views* e *Controllers*) simplificada, tendo em vista uma clara separação das responsabilidades e promovendo a autonomia das camadas em um projecto C#. Esta abordagem estrutural não só facilita a manutenção e a evolução do código, mas também sustenta a integração e a colaboração eficaz entre as diferentes partes da aplicação. A seguir, detalha-se a disposição das pastas que compõem a aplicação, cada uma desempenhando um papel específico dentro do ecossistema de *software*:

-   `Views (Apresentação)/`: interface do utilizador (formulário de artigos) que permite a interação com o sistema.
    -   `UIFichaArtigos.cs:`: formulário para exibir e editar os dados de um artigo. Contém a aba "Stocks" e nela é disparado o evento para chamar o método ``` ActualizaCustoPadrao() ``` contido no ficheiro ``` ArtigoController ```.
-   `Controllers (Controlador)/`: lida com as requisições do utilizador, coordena as acções entre o Modelo (*Models*) e a Apresentação (*Views*) e contém a lógica de controlo da aplicação.
    -   `ArtigoController.cs`: contém o método ``` void ActualizarCustoPadrao(BasBEArtigo codigoArtigo) ``` para actualizar o custo padrão de um artigo, que utiliza o preço de custo da última compra e a respectiva taxa de câmbio efectua uma actualização à base de dados.
-   `Models (Modelo)/`: representa os dados e as regras de negócio da aplicação. Inclui as classes que representam os objectos do sistema (Artigo) e as interfaces e classes de acesso aos dados fornecidos pela **Plataforma Cegid Primavera** (BasBEArtigo).
    -   `ArtigoRepository.cs`: classe que implementa a classe extensível **(BasBEArtigoFornecedor)** e contém o método ``` DateTime GetDataUltimaCompra(BasBEArtigo codArtigo) ``` para obter a data da última compra de um determinado artigo.
    -   `ComprasRepository.cs`: classe que contém os métodos para obter o preço de custo da última compra, a moeda utilizada na última compra e a taxa de câmbio praticada no momento da última compra de um determinado artigo.
    -   `ConexaoBD.cs`: contém as variáveis de conexão à base de dados.
-   `Exceptions (Excepções)/`: responsável pelo tratamento de erros e excepções que ocorrem durante a execução da aplicação.
    -   `ErrorMessagePaneType.cs`: classe que implementa a classe extensível **(BasBSArtigos)** e contém o método ``` void DisplayError(String mensagemErro) ``` que devolve um tipo específico de "***MessageBox***" customizada herdada a partir da classe "BasBSArtigos".
    -   `ExceptionHandler.cs`: classe que implementa a super-classe **(*Exception*)** e que contém o método ``` void MostrarMensagemErro(String mensagemErro) ``` que devolve uma excepção em caso de ocorrência de erros durante a execução da aplicação.
    -   `FileLogger.cs`: classe que permite a criação de um ficheiro de log, sempre que ocorrer um erro ou excepção durante a execução da aplicação.
-   `Tests (Testes)/`: pasta que contém classes de testes unitários e de integração que garantem a qualidade do código e o funcionamento correcto da aplicação.

## 📦 Implantação

Compilar o código desenvolvido, que é convertido para um arquivo do tipo **dll** e para testar, devemos aceder à aplicação Primavera e configurar a extensão criada, recorrendo ao caminho "\bin\Debug", dentro do projecto.

![pri-extensibilidade](/CegidPrimaveraExtensibilitySamp1/Base/Assets/img/pri_extensib.png "Extension")

Para a implantação da *script* na base de dados que sustenta a aplicação Cegid Primavera ERP v10, será necessário construir, primeiramente, a instrução de consulta (***query***), em **linguagem T-SQL**, e para tal, será utilizada a ferramenta *Microsoft SQL Server Management Studio*. 

Na aplicação MSSMS, o *script* que se pretende criar será disparado por evento, a cada inserção de novas informações à base de dados. Esse evento é conhecido como ***trigger*** ou **gatilho** em português.

Um *trigger* permite disparar um evento a cada inserção (***INSERT***) ou alteração (***UPDATE***) de informações em uma respectiva tabela, na base de dados. Eis o exemplo da criação de um *trigger*, para ambos os casos supracitados:

``` sql

        -- Cria um evento que é disparado após a inserção de uma nova venda

        CREATE TRIGGER trg_CriaLogTransaccao
        ON TabelaVendas
        AFTER INSERT
        AS
        BEGIN
        SET NOCOUNT ON;
    
        UPDATE TabelaVendas
        SET TabelaVendas.DataGravacaoLog = inserted.DataVenda
        FROM TabelaVendas
        INNER JOIN inserted ON TabelaVendas.ID = inserted.ID
        WHERE TabelaVendas.TipoFactura != 'ORÇAMENTO'
        END
        
        GO
```
``` sql

        -- Cria um evento que é disparado após a alteração de um produto

        CREATE TRIGGER trg_AlteraLog
        ON TabelaProdutos
        AFTER UPDATE
        AS
        BEGIN
        SET NOCOUNT ON;
    
        UPDATE TabelaProdutos
        SET TabelaProdutos.DataUltimaAlteracao = updated.DataGravacao
        FROM TabelaProdutos
        INNER JOIN updated ON TabelaProdutos.ID = updated.ID
        WHERE TabelaProdutos.Item LIKE 'COMPUTADOR%'
        END
        
        GO
```

Após à criação do *trigger*, deve-se aceder à aplicação Cegid Primavera ERP v10, **no módulo de Vendas**, na secção **"Pagamentos e Recebimentos"**, conforme demonstra a figura abaixo.

![pri-section](/assets/img/pri_section.png "Payables and Receivables Section")

Depois, deve-se clicar em **"Exploração"**, seguido do item de menu **"Mais"** e escolher a opção **"Documentos Emitidos"**, conforme demonstram as figuras abaixo.

![pri-section](/assets/img/pri_explore.png "Explore")
![pri-item](/assets/img/pri_item.png "See More")

Será exibido um formulário, no qual deve-se dar os seguintes passos:

- Seleccionar o **tipo de destinatário** (Clientes, Fornecedores, etc), seleccionar as **datas** para a filtragem de informações, seleccionar o **tipo de documentos** a consultar e as restantes opções, que são facultativas; conforme demonstra a figura abaixo.

  ![pri-form](/assets/img/pri_form.png "Issued Documents")

- A seguir, deve-se clicar em **"Imprimir"**, para que seja emitido o relatório, conforme demonstra a figura abaixo.

  ![pri-report](/assets/img/pri_report.png "Issued Documents Report")

## 🏆 Resultado obtido

De acordo ao plano de projecto, o objectivo foi atingido. Com a criação do *script*, através do qual foi gerado o *trigger*, foi possível actualizar a base de dados com as informações solicitadas pelo cliente e foi criado um evento que permitirá que esta actualização seja sempre efectuada, sempre que houver uma inserção de novas informações, na base de dados. Veja as figuras abaixo.

Inicialmente, algumas informações não podiam ser visualizadas pelo cliente, conforme demonstra a figura abaixo.

![pri-objective](/assets/img/pri_objective1.png "Report With No Data")

Após às devidas alterações, as informações já podem ser visualizadas pelo cliente, conforme demonstra a figura abaixo.

![pri-objective](/assets/img/pri_objective2.png "Report With Required Data")

## 📄 Licença

Este projecto está sob a licença do MIT - veja o arquivo [LICENSE](https://github.com/emjit-lda/Cegid-Primavera-ERP-v10-Extensibility-SAMP5/blob/shika-dev/LICENSE) para detalhes.

Apenas reserva-se a sua alteração e difusão sob responsabilidade da **[EMJIT, LDA](https://www.emjit.com)**.

## 🛠 Colaboradores

* **@[Stéfano Lorenzo](https://github.com/Stefano-Lorenzo)**.


## 👁 Visitas

<h3>
    <p align="center">
      <br>
      <img align="center" src="https://profile-counter.glitch.me/emjit-lda/count.svg" />
    </p>
</h3>

⌨️ Copyright © 2025 **[EMJIT, LDA](https://www.emjit.com)** - Consultoria e Soluções Tecnológicas ![emjit-02](/assets/img/emjit02.png "EMJIT, LDA")