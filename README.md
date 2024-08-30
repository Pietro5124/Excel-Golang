# Excel-Golang

Este repositório é um exemplo prático de como abrir um arquivo Excel na linguagem Golang.

# Leitura de Arquivo Excel em Go usando github.com/xuri/excelize/v2

Neste exemplo, mostraremos como ler um arquivo Excel usando a biblioteca `github.com/xuri/excelize/v2` na linguagem de programação Go. Essa biblioteca fornece funcionalidades avançadas para lidar com arquivos Excel.

## Instalação da Biblioteca

Antes de começar, você precisa instalar a biblioteca `excelize/v2`. Siga os passos abaixo:

1. Abra o terminal do seu projeto Go.
2. Execute o comando para instalar a biblioteca:

    ```sh
    go get github.com/xuri/excelize/v2
    ```

## Definindo a Struct

Primeiro, você deve alterar a struct `linha`, que representará os dados que serão lidos do arquivo Excel. Ela está localizada no arquivo `struct.go`:

1. Crie a struct `linha` no seu código Go:

    ```go
    type linha struct {
        Conta      int       `xlsx:"Conta"`
        Nome       string    `xlsx:"NOME"`
        Quantidade int       `xlsx:"QUANTIDADE"`
        Data       time.Time `xlsx:"DATA"`
        Escritorio string    `xlsx:"ESCRITÓRIO"`
    }
    ```

2. **Nota Importante:** Para que os campos da struct sejam corretamente mapeados, o nome definido na tag `xlsx` deve ser exatamente igual ao nome da coluna no arquivo Excel.

## Leitura do Arquivo Excel

Com a struct definida, agora você pode ler os dados do arquivo Excel:

1. Primeiro, você precisa alterar a variável `filePath` no início do arquivo `app.go` para o caminho correto do arquivo Excel na sua máquina:

    ```go
    // Caminho do arquivo
    var filePath string = "caminho/do/seu/arquivo/exemplo.xlsx"
    ```

2. Segundo, caso seja necessário, altere a variável `starLine`, que define a linha inicial de leitura do Excel. Essa variável também é inicializada no início do arquivo `app.go`. **Observação importante:** A linha de início no Excel deve conter os nomes das colunas.

    ```go
    // Linha de início que o Excel começa a ler (a contagem de linhas começa em 0)
    var starLine int = 0
    ```

3. Caso a planilha no Excel tenha um nome diferente, altere a variável `tableName`. Essa é a variável que define a planilha que será lida no Excel. Essa variável também é inicializada no início do arquivo `app.go`:

    ```go
    var tableName string = "Planilha1"
    ```

4. **Executando o Código:** Antes de executar o código, certifique-se de que todas as variáveis estejam configuradas corretamente e que o arquivo exista no local especificado. Execute o código:

    ```sh
    go run app.go
    ```

## Pontos Importantes

- A formatação da data deve ser "yyyy-mm-dd". Caso queira utilizar outro formato, altere a função `parseTime` para o modelo desejado.
- Os tipos definidos na struct são muito importantes para a formatação correta dos dados lidos do arquivo.
- Quando executar o programa, você deve garantir que o arquivo Excel está no caminho correto, senão ocorrerá um erro de acesso ao arquivo.
- Se o arquivo Excel contiver várias planilhas, você pode precisar iterar sobre elas ou especificar a planilha correta ao ler os dados. Use `GetSheetList` para obter uma lista de planilhas disponíveis no arquivo:

    ```go
    sheets := file.GetSheetList()
    for _, sheet := range sheets {
        fmt.Println(sheet)
    }
    ```

- Valide os dados lidos do Excel antes de processá-los. Certifique-se de que os valores atendem às expectativas (por exemplo, números estão no intervalo esperado, datas estão formatadas corretamente, etc.).
