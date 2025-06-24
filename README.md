# Simulador de Investimentos em Fundos Imobiliários (FIIs) em Excel

Bem-vindo ao repositório do Simulador de Investimentos em FIIs! Esta ferramenta foi desenvolvida em Microsoft Excel para ser uma solução completa e intuitiva, permitindo ao usuário projetar o crescimento de seu patrimônio, simular cenários de longo prazo e diversificar sua carteira de investimentos com base em seu perfil de risco.

<p align="center">
  <img src="./images/image_3fde1c.png" alt="Visão Geral do Simulador" width="600">
</p>

## Visão Geral do Projeto

O objetivo deste projeto é ir além de uma simples planilha de cálculos. Trata-se de uma **ferramenta de análise financeira** que responde a perguntas de negócio essenciais para qualquer investidor de fundos imobiliários:

* Quanto patrimônio terei acumulado investindo um valor X por Y anos?
* Qual será minha renda passiva mensal (dividendos) no futuro?
* Como meu patrimônio evolui em diferentes prazos (2, 5, 10, 20, 30 anos)?
* Como devo distribuir meus aportes mensais de acordo com meu perfil de investidor (Conservador, Moderado, Agressivo)?

Para entregar uma experiência de alta qualidade, a ferramenta foi construída sobre três pilares: **lógica financeira sólida**, **boas práticas de desenvolvimento em Excel** e uma **interface de usuário limpa e focada**, que simula a usabilidade de um aplicativo desktop.

## Estrutura e Desenvolvimento Técnico

A versão final do projeto implementa técnicas avançadas para garantir robustez, escalabilidade e uma excelente experiência de uso. A seguir, detalhamos a arquitetura da solução.

### 1. Central de Dados com a Planilha `Configuracoes`

Toda a lógica de parametrização foi centralizada em uma planilha de apoio chamada `Configuracoes`. Essa abordagem desacopla os dados da interface principal, tornando a manutenção e a expansão do projeto muito mais simples.

Nesta planilha, foram criadas **Tabelas Nomeadas** do Excel (`tab_perfil`, `tab_fii`, `tab_cenarios`, `tab_chave`), que servem como fontes de dados dinâmicas para o simulador.

### 2. Validação de Dados Dinâmica com `INDIRETO`

Para os campos de seleção do usuário (Perfil, Cenários), em vez de listas estáticas, utilizamos a função `INDIRETO` combinada com as Tabelas Nomeadas.

* **Seleção de Perfil:** `=INDIRETO("tab_perfil[PERFIL]")`
* **Seleção de Cenários:** `=INDIRETO("tab_cenarios[CENÁRIOS]")`

Essa técnica permite que as listas de opções sejam atualizadas automaticamente caso novos perfis ou cenários sejam adicionados à tabela na planilha `Configuracoes`, sem a necessidade de editar a validação de dados manualmente.

### 3. Lógica de Alocação com `PROCX`

Para determinar o percentual de alocação sugerido, a função `PROCV` foi substituída pela mais moderna e poderosa **`PROCX`**. A fórmula cria uma chave de busca composta, concatenando o perfil selecionado com o tipo de FII, e busca essa chave na tabela `tab_chave`.

**Fórmula utilizada:** `=PROCX(perfil&"-"&$I34;tab_chave[CHAVE];tab_chave[%];"";0)`

* `perfil&"-"&$I34`: Cria a chave de busca (ex: "CONSERVADOR-TIJOLO").
* `tab_chave[CHAVE]`: Coluna onde a chave será procurada.
* `tab_chave[%]`: Coluna da qual o resultado (a porcentagem) será retornado.

Esta abordagem é mais eficiente e flexível que o `PROCV`, pois não depende da ordem das colunas.

### 4. Interface e Experiência do Usuário (UX)

A interface foi projetada para ser intuitiva e autoexplicativa.

* **Comentários Guiados:** Foram adicionados comentários explicativos nas células de input principais, orientando o usuário sobre como preencher os dados.
* **Botão de Edição (Modo Desenvolvedor):** Para facilitar a manutenção da planilha, foi criado um mecanismo que permite ao desenvolvedor sair do "modo aplicativo" e retornar à visualização padrão do Excel.

## Configuração do Ambiente (VBA)

Para criar a experiência de "aplicativo", utilizamos um código VBA que é executado automaticamente ao abrir e fechar a pasta de trabalho.

### 1. Código para o Módulo `ThisWorkbook`

*(Este script gerencia os eventos de abrir e fechar a pasta de trabalho)*

```vb
' =========================================================================
' CÓDIGO A SER INSERIDO NO MÓDULO DE OBJETO "ThisWorkbook"
' =========================================================================
' Este código executa as macros de configuração e restauração automaticamente
' ao abrir e fechar a pasta de trabalho.

Option Explicit

Private Sub Workbook_Open()
    ' Chama a rotina principal para configurar a visualização de "App".
    Call ConfigurarVisualizacaoApp
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Chama a rotina que restaura as configurações padrão do Excel.
    ' Isso é MUITO IMPORTANTE para não afetar outras planilhas que o usuário abrir.
    Call RestaurarVisualizacaoPadrao
End Sub
```

### 2. Código para um Módulo Padrão (ex: `Módulo1`)

*(Este script contém a lógica principal da aplicação)*

```vb
' =========================================================================
' CÓDIGO A SER INSERIDO EM UM NOVO MÓDULO PADRÃO (EX: Módulo1)
' =========================================================================
' Este módulo contém a lógica principal para alterar e restaurar a aparência do Excel.

Option Explicit

' Defina aqui a sua senha de proteção. Deixe em branco ("") para não usar senha por padrão.
Private Const SENHA_PROTECAO As String = ""

Public Sub ConfigurarVisualizacaoApp()
    ' Macro principal para configurar a planilha com aparência de aplicativo.
    Dim ws As Worksheet
    
    ' ATENÇÃO: Altere "Invest" para o nome exato da sua planilha principal.
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Invest")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "A planilha de nome 'Invest' não foi encontrada." & vbCrLf & "A macro de inicialização não pode ser executada.", vbCritical, "Erro de Configuração"
        Exit Sub
    End If
    
    ws.Activate
    
    ' Oculta elementos gerais da aplicação Excel.
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    
    ' Configura a janela ativa para o modo "App".
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayWorkbookTabs = False
        .Zoom = 100
    End With
    
    ' Define a área máxima de rolagem.
    ws.ScrollArea = "A1:Q60"
    
    Application.Goto Reference:=ws.Range("A1"), Scroll:=True
    
    ' Protege a planilha, permitindo a execução de macros.
    ws.Protect Password:=SENHA_PROTECAO, UserInterfaceOnly:=True
End Sub

Public Sub RestaurarVisualizacaoPadrao()
    ' Macro para restaurar as configurações padrão do Excel para edição.
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Invest")
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    ' Desprotege a planilha
    ws.Unprotect Password:=SENHA_PROTECAO

    ' Remove a limitação da área de rolagem.
    ws.ScrollArea = ""

    ' Restaura os elementos da aplicação.
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    
    ' Restaura a visualização da janela.
    With ActiveWindow
        .DisplayGridlines = False ' Mantém as linhas de grade ocultas.
        .DisplayHeadings = True
        .DisplayWorkbookTabs = True
    End With
End Sub
```

### 4.1. Botão de Acesso ao Modo de Edição

Para facilitar a edição e manutenção da planilha, foi inserido um ícone de seta que funciona como um **botão de acesso ao modo de edição**. Ao ser clicado, ele executa a macro `RestaurarVisualizacaoPadrao`, que desativa a interface de "aplicativo" e reexibe as ferramentas padrão do Excel.

![Ícone para alternar para o Modo de Edição](./images/seta.png)

Para replicar essa funcionalidade em seu projeto, siga os passos:

1.  **Inserir um Ícone ou Forma:**
    * Na guia `Inserir` > `Ilustrações`, escolha uma `Forma` ou um `Ícone` de sua preferência (no projeto, foi utilizada a seta verde).
    * Posicione o objeto na planilha em um local de fácil acesso.

2.  **Atribuir a Macro `RestaurarVisualizacaoPadrao`:**
    * Clique com o **botão direito** no objeto inserido.
    * No menu de contexto, selecione **`Atribuir Macro...`**.
    * Na janela que se abre, selecione a macro **`RestaurarVisualizacaoPadrao`** e clique em `OK`.

> **Dica de Desenvolvedor:** Se precisar editar a planilha mas a macro de inicialização já a travou, feche e reabra o arquivo mantendo a tecla **`SHIFT` pressionada**. Isso impedirá a execução automática da macro `Workbook_Open`, dando-lhe acesso total à planilha.

## Créditos e Desenvolvimento

Este projeto, sua lógica, estrutura de planilha foram desenvolvidas sem auxílio de IA. 
A documentação e o código VBA, foi desenvolvido com o auxílio da IA **Gemini Pro**, da Google. A ferramenta foi utilizada como um assistente de desenvolvimento para acelerar a criação do código, estruturar os conceitos e gerar esta documentação.
