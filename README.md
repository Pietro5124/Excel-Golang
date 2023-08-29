# Excel-Golang
Esse repositorio é um exemplo pratico de como abrir um arquivo excel na linguagem Golang
# Leitura de Arquivo Excel em Go usando github.com/xuri/excelize/v2

Neste exemplo, mostraremos como ler um arquivo Excel usando a biblioteca `github.com/xuri/excelize/v2` na linguagem de programação Go. Essa biblioteca fornece funcionalidades avançadas para lidar com arquivos Excel.

## Instalação da Biblioteca

Antes de começar, você precisa instalar a biblioteca `excelize/v2`. Você pode fazer isso usando o seguinte comando:

```sh
go get github.com/xuri/excelize/v2
```
Definindo a Struct
Primeiro, vamos definir uma struct que representará os dados que serão lidos do arquivo Excel:

```sh
type Linha struct {
	Conta      int       `xlsx:"Conta"`
	Nome       string    `xlsx:"NOME"`
	Quantidade int       `xlsx:"QUANTIDADE"`
	Data       time.Time `xlsx:"DATA"`
	Escritorio string    `xlsx:"ESCRITÓRIO"`
}
```
Para vincular o campo da struct o nome que  definido no campo xlsx tem que ser o nome igual da coluna da tabela.