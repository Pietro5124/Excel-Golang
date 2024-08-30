package main

import (
	"fmt"
	"log"
	"main/src/excel"
	"os"
)

// Path do arquivo
var filePath string = "Exemplo.xlsx"

// Linha de Inicio que o excel começa a ler obs: A lista inicia a contagem de linha pelo 0
var starLine int = 0

// Nome da planilha do Excel que vai ser lida do arquivo
var tableName string = "Planilha1"

func main() {

	// Abre o arquivo para leitura
	file, err := os.Open(filePath)
	if err != nil {
		// Campo "file" não está presente no formulário
		log.Panicln("Esse arquivo não existe:", err.Error())
		return
	}
	defer file.Close()
	lines, err := excel.OpenExcel(file, starLine, tableName)
	if err != nil {
		// Campo "file" não está presente no formulário
		log.Panicln("Erro na converção em struct", err.Error())
		return
	}
	for _, line := range lines {
		// Neste ponto, você já pode utilizar as informações. Se ocorrer algum erro, consulte os pontos importantes no arquivo.
		// Mostra as linhas da tabela
		fmt.Println(" Nome:", line.Nome, " Conta:", line.Conta, " Data:", line.Data, " Quantidade:", line.Quantidade, " Telefone:", line.Telefone)
	}

}
