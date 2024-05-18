# Curso VBA

## Aula 1

1. Introdução - Aula 1
   1. Ambiente de desenvolvimento
   2. Acessar a aba de desenvolvedor

2. Linguagem de progamação e algoritmos
   1. Variáveis
   2. Vetores
   3. Listas
   4. Dicionarios
   5. Estruturas Condicionais
   6. Estruturas de Repetição
   7. Funções
   8. Classes

## Aula 2

1. Macros
   1. Introdução aos Macros
   2. Gravar Macros
   3. Editar Macros
   4. Executar Macros
2. Funções VBA
   1. Funções Embutidas
   2. Funções Personalizadas

## Aula 3

1. Manipulando planilhas
2. Ambiente Gráfico
3. Manipulação de arquivos
   1. Ler arquivo TXT
   2. Escrever Arquivo TXT

## Aula 4

1. Acessar dados Externos
   1. Acessar dados por API
   2. Obter dados via importação
2. Definição e elaboração de um projeto

## Aula 5

Continuação da eleboração do projeto e apresentação.

## Introdução ao VBA

O Visual Basic for Applications (VBA) é uma linguagem de programação poderosa e versátil integrada ao Microsoft Office. Ele permite automatizar tarefas, criar macros e interfaces de usuário personalizadas, estender as funcionalidades dos aplicativos do Office e muito mais. O VBA é relativamente fácil de aprender, especialmente para quem já tem conhecimento de outras linguagens de programação.

Benefícios do VBA:

Automação de Tarefas: Automatize tarefas repetitivas e complexas, economizando tempo e esforço.
Criação de Macros: Crie macros para executar sequências de comandos com apenas um clique.
Desenvolvimento de Interfaces: Crie interfaces de usuário personalizadas e interativas para facilitar o uso das aplicações.
Extensão de Funcionalidades: Amplie as funcionalidades dos aplicativos do Office com recursos personalizados.
Integração com Outros Sistemas: Integre o VBA com outros sistemas e bancos de dados para criar soluções mais robustas.
Aplicações do VBA:

Automatização de Tarefas Administrativas: Crie macros para formatar relatórios, enviar emails, gerar gráficos e muito mais.
Desenvolvimento de Ferramentas de Negócios: Crie ferramentas personalizadas para análise de dados, gerenciamento de projetos e outras necessidades de negócios.
Criação de Aplicativos Educacionais: Crie jogos educativos, simulações interativas e outros recursos para aprimorar o aprendizado.
Desenvolvimento de Jogos: Crie jogos simples com gráficos e animações usando o VBA.
Desenvolvimento de Sites Dinâmicos: Utilize o VBA para interagir com sites dinâmicos e bancos de dados.
2. Ambiente de Desenvolvimento

O VBA é desenvolvido no ambiente de desenvolvimento integrado (IDE) do Microsoft Office, que oferece diversas ferramentas para facilitar a escrita, depuração e execução de código.

Acessando a Aba Desenvolvedor:

Abra um aplicativo do Office (por exemplo, Excel, Word).
Clique na guia "Arquivo".
Selecione "Opções".
Na janela "Opções", clique em "Personalizar Faixa de Opções".
Na caixa "Principal", marque a caixa de seleção "Desenvolvedor".
Clique em "OK".
A guia "Desenvolvedor" agora estará disponível na faixa de opções, fornecendo acesso às ferramentas de desenvolvimento do VBA.

1. Linguagem de Programação e Algoritmos

A linguagem de programação VBA é composta por comandos, funções e estruturas de controle que permitem realizar diversas operações. Para escrever código VBA eficaz, é importante entender os conceitos básicos de programação, como variáveis, vetores, listas, dicionários, estruturas condicionais, estruturas de repetição, funções e classes.

1.1 Variáveis:

As variáveis armazenam dados que podem ser utilizados em diferentes partes do código. Cada variável possui um nome e um tipo de dado (texto, número, data, etc.).

Exemplos de Declaração de Variáveis:

~~~vba
Dim nome As String
Dim idade As Integer
Dim salario As Double
Dim dataNascimento As Date
~~~

1.2 Vetores:

Os vetores armazenam coleções de dados do mesmo tipo. Cada elemento do vetor possui um índice que permite acessá-lo individualmente.

Exemplos de Declaração e Manipulação de Vetores:

~~~VBA
Dim notas(10) As Integer
notas(1) = 90
notas(2) = 85
notas(3) = 78

Dim media As Integer
For i = 1 To 10
  media = media + notas(i)
Next i

MsgBox "Média: " & media
~~~

1.3 Listas:

As listas armazenam coleções de dados de qualquer tipo. São mais flexíveis que os vetores, pois permitem adicionar, remover e modificar elementos dinamicamente.

Exemplos de Declaração e Manipulação de Listas:

~~~VBA
Dim alunos As New Collection
alunos.Add "João"
alunos.Add "Maria"
alunos.Add "Pedro"

MsgBox "Quantidade de alunos: " & alunos.Count

For Each aluno In alunos
  MsgBox aluno
Next aluno
~~~

1.4 Dicionários:

Os dicionários armazenam coleções de dados em pares chave-valor. As chaves servem como identificadores únicos para acessar os valores correspondentes.

Exemplos de Declaração e Manipulação de Dicionários:

~~~VBA
Dim cidades As New Dictionary
cidades.Add "SP", "São Paulo"
cidades
~~~

1.5 Estruturas Condicionais:

Estruturas condicionais permitem tomar decisões no código com base em condições específicas. As principais estruturas condicionais no VBA são:

IF...THEN...ELSE: Verifica se uma condição é verdadeira e executa um bloco de código se for, ou outro bloco se não for.
SWITCH...CASE: Permite escolher entre várias ações com base em diferentes valores de uma variável.
SELECT CASE: Similar ao SWITCH...CASE, mas com opções mais flexíveis para comparar valores.

1.5.1 IF

A estrutura IF...THEN...ELSE é a mais utilizada e oferece flexibilidade para lidar com diversas situações. Funciona da seguinte maneira:

IF (condição): Verifica se a condição é verdadeira.
THEN: Se a condição for verdadeira, o bloco de código THEN será executado.
ELSE: Se a condição for falsa, o bloco de código ELSE será executado.

Exemplo:

~~~vba
' Verificando se o usuário é maior de idade
Dim idade As Integer
idade = InputBox("Digite sua idade:")

If idade >= 18 Then
    ' Usuário é maior de idade
    MsgBox "Você pode entrar na festa!"
Else
    ' Usuário é menor de idade
    MsgBox "Você não pode entrar na festa."
End If
~~~

1.5.2 SWITCH...CASE

Quando você precisa lidar com diversas possibilidades e tomar ações específicas para cada uma, o SWITCH...CASE é a escolha ideal. Ele funciona como um menu de opções:

* SWITCH (variável): Define a variável que será usada para a comparação.
* CASE (valor1): Bloco de código executado se a variável for igual ao valor1.
* CASE (valor2): Bloco de código executado se a variável for igual ao valor2.
... outros casos: Mais blocos CASE para outros valores.
* DEFAULT: Bloco de código executado se nenhum dos valores CASE corresponder.

Exemplo 1:

~~~vba
' Selecionando o dia da semana e exibindo uma mensagem personalizada
Dim diaSemana As String
diaSemana = InputBox("Digite o dia da semana (segunda, terça, ...):")

Select Case diaSemana
    Case "segunda"
        MsgBox "Segunda-feira: Hora de começar a semana com foco!"
    Case "terça"
        MsgBox "Terça-feira: Superando o hump day com energia!"
    ' ... outros casos para os demais dias da semana
    Default
        MsgBox "Dia da semana inválido. Tente novamente."
End Select
~~~

Exemplo 2:

~~~vba
' Selecionando o dia da semana e exibindo uma mensagem personalizada
Dim diaSemana As String
diaSemana = InputBox("Digite o dia da semana (segunda, terça, ...):")

Select Case diaSemana
    Case "segunda"
        MsgBox "Segunda-feira: Hora de começar a semana com foco!"
    Case "terça"
        MsgBox "Terça-feira: Superando o hump day com energia!"
    ' ... outros casos para os demais dias da semana
    Default
        MsgBox "Dia da semana inválido. Tente novamente."
End Select
~~~

1.6 Estruturas de Repetição:

1.6.1 FOR...NEXT: Repetindo um Número Fixo de Vezes
A estrutura FOR...NEXT é perfeita quando você precisa executar um bloco de código um número específico de vezes. Funciona da seguinte maneira:

FOR (contador) = (valor inicial) TO (valor final): Define o contador e o intervalo de repetições.
NEXT (contador): Incrementa o contador e verifica se ainda está dentro do intervalo.

~~~vba
' Imprimindo números de 1 a 10
For i = 1 To 10
    MsgBox i
Next i
~~~

1.6.2  DO...WHILE: Repetindo Enquanto Uma Condição For Verdadeira
Com o DO...WHILE, você repete o código enquanto uma condição específica permanecer verdadeira. É ideal para situações em que você não sabe quantas vezes o loop será executado.

DO While (condição): Verifica se a condição é verdadeira.
LOOP: Executa o bloco de código enquanto a condição for verdadeira.

~~~vba
' Calculando a soma de números ímpares até 100
Dim soma As Integer
soma = 0
i = 1

Do While i <= 100
    If i Mod 2 <> 0 Then ' Verifica se o número é ímpar
        soma = soma + i
    End If
    i = i + 1
Loop

MsgBox "Soma dos números ímpares até 100: " & soma
~~~

1.6.3 FOR EACH: Repetindo em Coleções de Dados

O FOR EACH facilita a iteração sobre cada elemento em uma coleção de dados, como arrays ou listas. É ideal para processar cada item individualmente.

FOR EACH (variável) IN (coleção): Define a variável que receberá cada elemento da coleção.
NEXT (variável): Passa para o próximo elemento da coleção.

~~~vba
' Imprimindo os nomes das capitais do Brasil
Dim estados As New Dictionary
estados.Add "AC", "Rio Branco"
estados.Add "AL", "Maceió"
' ... outros estados e capitais

For Each estado In estados
    MsgBox estado & " - " & estados(estado)
Next estado
~~~

1.7 Funções

As funções no VBA são como blocos de construção mágicos que permitem modularizar seu código, torná-lo mais reutilizável, legível e eficiente. Com elas, você encapsula tarefas específicas em unidades independentes, facilitando a organização e a manutenção de seus projetos.

ara criar uma função, utilize a palavra-chave Function seguida do nome da função, seus parâmetros (entre parênteses) e o tipo de valor de retorno (opcional). O bloco de código da função fica entre Function e End Function.

~~~vba
' Função que calcula a soma de dois números
Function somar(numero1 As Integer, numero2 As Integer) As Integer
    somar = numero1 + numero2
End Function

' Chamando a função para somar 5 e 10
Dim resultado As Integer
resultado = somar(5, 10)
MsgBox "Resultado da soma: " & resultado
~~~

Funções sem Retorno:

Se a função não retornar um valor, utilize Sub no lugar de Function e omita As Integer ou o tipo de retorno.

1.8 Classes

No VBA, as classes são como plantas mágicas que geram objetos poderosos e reutilizáveis. Através da Programação Orientada a Objetos (POO), você organiza seu código em módulos distintos, encapsulando dados e comportamentos, promovendo a modularidade, flexibilidade e reutilização em seus projetos.

1.8.1 Criando Classes e Objetos:

Para criar uma classe, utilize a palavra-chave Class seguida do nome da classe e End Class. Dentro da classe, você define propriedades (dados) e métodos (comportamentos) que pertencem àquela classe.

~~~vba
' Classe para representar um cliente
Class Cliente

    ' Propriedades (dados)
    Private nome As String
    Private email As String
    Private telefone As String

    ' Métodos (comportamentos)
    Public Sub NovoCliente(nomeCliente As String, emailCliente As String, telefoneCliente As String)
        nome = nomeCliente
        email = emailCliente
        telefone = telefoneCliente
    End Sub

    Public Function ObterNome() As String
        ObterNome = nome
    End Function

    Public Function ObterEmail() As String
        ObterEmail = email
    End Function

    Public Function ObterTelefone() As String
        ObterTelefone = telefone
    End Function

End Class
~~~

Após criar a classe, você instancia objetos a partir dela utilizando a palavra-chave New. Cada objeto possui suas próprias propriedades e métodos, acessíveis através do ponto (.).

~~~vba
' Criando um objeto cliente
Dim cliente1 As New Cliente
cliente1.NovoCliente "João Silva", "joaosilva@email.com", "(11) 9999-9999"

' Acessando propriedades do objeto
MsgBox cliente1.ObterNome() ' Exibe "João Silva"
MsgBox cliente1.ObterEmail() ' Exibe "joaosilva@email.com"
MsgBox cliente1.ObterTelefone() ' Exibe "(11) 9999-9999"
~~~
