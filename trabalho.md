# Estrutura de linguagem #

1. Estrutura da linguagem
    * Variáveis:
    Em Python as variáveis são dinâmicamente tipadas ou seja no momento de sua declaração ou reatribuição o interpretador aufere o tipo.

      ~~~python
        nome = "Fulano de Tal"
        idade = 30
        pi = 3.14159
        print(nome)
        print(idade)
        print(pi)
      ~~~

      Essa forma de euferencia de tipos dá uma dinamicidade e facilita o desenvolvedor, deixando o código menos verboso, porém possibilita a mudança de tipo em tempo de execução.

      No exemplo abaixo uma variável troca de tipo apenas atribuindo um tipo diferente do inicial.

    ~~~python
    nome="willie" ## a variavel nome neste estágio é uma String, por auferencia
    nome = 123 ## Com esta reatribuição a variável nome passa a ser um inteiro
    ~~~

    Observação: O Python não possuí constantes.
    * Vetores:

        Os vetores, também conhecidos como arrays, permitem armazenar uma coleção de valores do mesmo tipo. Em Python, vetores são representados por listas, utilizando colchetes ([]) para definir seus elementos.

        ~~~python
        frutas = ["maçã", "banana", "laranja"]
        numeros = [1, 2, 3, 4, 5]
        temperaturas = [25.2, 22.8, 31.4]
        print(frutas[0])  # Acessando o primeiro elemento (índice 0)
        print(numeros[2])  # Acessando o terceiro elemento (índice 2)
        print(temperaturas[-1])  # Acessando o último elemento (índice -1)
        ~~~

    * Listas:
        As listas em Python são estruturas de dados versáteis, similares aos vetores, mas com a vantagem de poder armazenar elementos de diferentes tipos. A sintaxe para criar listas é a mesma dos vetores, utilizando colchetes ([]).

        ~~~python
        dados = ["Fulano", 30, 1.70, True]
        compras = ["leite", "ovos", "pão", "queijo"]
        misturado = [10, "Olá", 3.14, False]
        print(dados[1])  # Acessando o segundo elemento
        print(compras[2:4])  # Fatiamento da lista (do índice 2 ao 4)
        print(misturado[:])  # Copiando toda a lista
        ~~~

    * Dicionários

        Os dicionários são estruturas de dados que armazenam pares de chave-valor, permitindo a organização de dados de forma associativa. Em Python, dicionários são representados por chaves ({}), utilizando chaves únicas para cada valor.

        ~~~python
        pessoa = {"nome": "Fulano", "idade": 30, "profissao": "Programador"}
        agenda = {"segunda": "Reunião", "terça": "Consulta", "quarta": "Treino"}
        estoque = {"camiseta": 10, "calça": 15, "vestido": 8}
        print(pessoa["nome"])  # Acessando o valor da chave "nome"
        print(agenda["terça"])  # Acessando o valor da chave "terça"
        print(estoque.keys())  # Exibindo as chaves do dicionário
        ~~~

    * Estruturas Condicionais:
        As estruturas condicionais permitem controlar o fluxo do programa com base em determinadas condições. Em Python, as principais estruturas condicionais são if e elif.

        ~~~python
        idade = 12
        if idade < 18
            print("Voce é menor de idade")
        elif(idade == 18)
            print("Você tem 18 anos")
        else
            print("Você é maior de idade")
        ~~~

        * Expressão Condicional:
            Um dos recursos do Python para facilitar a vida do desenvolvedor é a expressão condicional que desempenha o mesmo papel da expressão condicional, porém com uma sintaxe reduzida.

            ~~~python
            idade = 12
            print("Maioridade" if idade >= 18 else "Menoridade") ##output "Menoridade"
            ~~~

    * Estruturas de repetição:
        O Python possui duas estruturas de repetição:
        * for:

            O "for" executa o trecho de código dentro de um intervalo.

            ~~~python
                for numero in range(1, 11):  # Repete 10 vezes
                    print(numero)
            ~~~

        * while:

            Executa o trecho de código enquanto condição for atendida.

            ~~~python
                while i < 5:  # Repete enquanto i for menor que 5
                    print("Valor atual:", i)
                    i += 1
            ~~~

    * Funções:
        Existem dois tipos de funções em Python, funções nomeadas e as funções não nomeadas (lambda).

        As funções nomeadas são comuns em outras linguagens, tem um nome como caracteristica, parâmetros e retorno.

        Uma caracteristica que deve ser ressaltada é que em função da linguagem não ser fortemente tipada, os parâmetros tambem não são tipados, podem levar a erros.

        Exemplo de função nomeada

        ~~~python
        def soma(x,y):
            resultado = x + y
            return resultado
        ~~~

        As funções não nomeadas (lamda) não recebem nome, por definição  são criadas, geralmente, para serem utilizadas por um período breve (dentro de um scopo de uma função, por exemplo).
        Elas são definidas da seguinte forma.

        lambda argumentos : expressão

        ~~~python
        media_1 = lambda valores : sum(valores) / len(valores)
        ~~~

    * Importação e uso de bibliotecas:

        Python oferece um vasto ecossistema de bibliotecas prontas para uso. Estas bibliotecas fornecem funções e classes para diversas tarefas, como manipulação de dados, matemática, plotagem e muito mais. Para importar uma biblioteca, utilizamos a instrução import.

        ~~~python
        import math
        # Utilizando a função sqrt() da biblioteca math
        raiz_quadrada = math.sqrt(16)
        print("A raiz quadrada de 16 é:", raiz_quadrada)
        ~~~

    * Classes

        As classes servem como modelos para a criação de objetos. Elas definem o comportamento e os atributos que os objetos de uma determinada classe irão possuir. Em Python, a criação de classes é feita utilizando a palavra-chave class.

        ~~~python
        class Pessoa:
        def __init__(self, nome, idade):
            self.nome = nome
            self.idade = idade

        def apresentacao(self):
            print("Olá, meu nome é", self.nome, "e tenho", self.idade, "anos.")

        # Criando objetos da classe Pessoa
        pessoa1 = Pessoa("Fulano", 30)
        pessoa2 = Pessoa("Beltrana", 25)

        pessoa1.apresentacao()
        pessoa2.apresentacao()
        ~~~

        * Metodos especiais:

            Para orientação a objetos a linguagem possui metódos específicos que podem ser criados ou sobrescritos na classe. Esse métodos tem como singularidade de serem declarados por dois underlines no inicio e dois underlines no fim do nome da classe.

           No exemplo anterior, o método __init__ da classe Pessoa recebe os argumentos nome e idade e os atribui aos atributos do objeto.
    * Encapsulamento e modificadores de acesso

        O encapsulamento é um dos pilares da programação orientada a objetos (POO), promovendo a organização, segurança e modularidade do código. Em Python, o encapsulamento é implementado de forma flexível, utilizando convenções e modificadores de acesso para controlar a visibilidade de atributos e métodos das classes.

        * Convenções de Encapsulamento

        Em Python, convenções de nomenclatura são utilizadas para indicar o nível de visibilidade dos atributos e métodos de uma classe:

        Público (_): Acessível de qualquer lugar. Ex: self.nome_publico
        Protegido (_): Acessível dentro da classe e subclasses. Ex: self._atributo_protegido
        Privado (__): Acessível apenas dentro da classe. Ex: self.__metodo_privado

        * Modificadores de Acesso

        Embora as convenções sejam widely used, Python oferece modificadores de acesso explícitos para maior controle:

        public: Torna o membro público (acessível de qualquer lugar).
        protected: Torna o membro protegido (acessível dentro da classe e subclasses).
        private: Torna o membro privado (acessível apenas dentro da classe).

        * Exemplo de Encapsulamento

        Considere a seguinte classe ContaCorrente:

        ~~~Python
        class ContaCorrente:
            def __init__(self, nome, saldo):
                self.nome = nome
                self.__saldo = saldo  # Atributo privado

            def get_saldo(self):
                return self.__saldo

            def depositar(self, valor):
                if valor > 0:
                    self.__saldo += valor
                else:
                    print("Valor inválido para depósito.")

            def sacar(self, valor):
                if valor > 0 and valor <= self.__saldo:
                    self.__saldo -= valor
                else:
                    print("Saldo insuficiente ou valor inválido para saque.")
        ~~~

    * Herança:

    A herança permite que uma classe (subclasse) herde atributos e métodos de outra classe (superclasse). Quando um método é chamado em um objeto de uma subclasse, o interpretador Python verifica primeiro se o método existe na subclasse. Se não existir, ele busca na superclasse e assim por diante na hierarquia de herança. Isso possibilita a criação de comportamentos especializados para subclasses, mantendo a implementação comum na superclasse.

    ~~~python
    class Forma:
        def __init__(self, cor):
        self.cor = cor

    def area(self):
        raise NotImplementedError("O método area() deve ser implementado na subclasse")

    class Retangulo(Forma):
    def __init__(self, cor, base, altura):
        super().__init__(cor)  # Chamando construtor da superclasse
        self.base = base
        self.altura = altura

    def area(self):
        return self.base * self.altura

    class Circulo(Forma):
    def __init__(self, cor, raio):
        super().__init__(cor)  # Chamando construtor da superclasse
        self.raio = raio

    def area(self):
        return 3.14 * (self.raio ** 2)

    # Criando objetos das subclasses
    retangulo = Retangulo("vermelho", 5, 4)
    circulo = Circulo("azul", 3)

    # Polimorfismo em ação: o mesmo método "area()" retorna resultados diferentes
    print("Área do retângulo:", retangulo.area())
    print("Área do círculo:", circulo.area())
    ~~~

* Classes Abstratas

    Em Python, classes abstratas servem como modelos para a criação de classes derivadas, definindo uma estrutura comum e funcionalidades básicas que devem ser implementadas nas subclasses. Elas não podem ser instanciadas diretamente, mas fornecem um roteiro para as classes filhas seguirem.

    Para que se possa utilizar classe abstratas na lingugem deve-se utilizar a biblioteca ABC e o decorator (abstractmethod).

    Exemplo de uso de classe abstrata:

    ~~~python
    from abc import ABC, abstractmethod

  class MinhaClasseAbstrata(ABC):
    @abstractmethod
    def escrever(self):
        print("Método abstrato escrever() chamado")
    @abstractmethod
    def outro_metodo_abstrato(self):
        pass
    ~~~

    Implementando uma classe abstrata:

    ~~~python
    class Minha_Classe(MinhaClasseAbstrata):
    def escrever(self):
        print("Método escrever() implementado")
    def outro_metodo_abstrato(self):
        print("Outro método abstrato implementado")
    ~~~

    Instanciando a classe "Minha_classe":

    ~~~python
    teste = Minha_Classe()
    teste.escrever() ##output M�todo escrever() implementado
    ~~~

    Protocols:

    Na linguagem Python não existem interfaces assim como outras linguagems orientadas a objeto, porém existe uma estrutura chamada Protocol que pode ser similiar a interface do C# por exemplo.

    O protocol exige que as classes implementem os métodos da classe que herda a  classe Protocol.

    ~~~python
    from typing import Protocol
    class Printable(Protocol):
    pages: int
    def print(self):
        pass

    class Book(Printable):
        def __init__(self, pages: int):
            self.pages = pages

    def print(self):
        print(f"Printing a book with {self.pages} pages.")

    book = Book(100)

    def print_object(obj: Printable):
        obj.print()

    print_object(book)
    ~~~

    Contudo, mesmo que Book, não implemente todos as definições indicadas dentro de Printable, o código não gera erro, ao menos na versão 3.10.8, utilizada em teste.
