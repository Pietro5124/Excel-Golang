package main

import (
	"fmt"
	"main/src/excel"
	"os"
)

// Path do arquivo
var filePath string = ""

func main() {

	// Abre o arquivo para leitura
	file, err := os.Open(filePath)
	if err != nil {
		// Campo "file" não está presente no formulário
		fmt.Println("Esse arquivo não existe")
		return
	}
	defer file.Close()
	lines, err := excel.OpenExcel(file)
	if err != nil {
		// Campo "file" não está presente no formulário
		fmt.Println("Erro na converção em struct")
		return
	}
	for _, line := range lines {
		// Mostra as linhas da tabela
		fmt.Println(line.Nome, line.Conta, line.Data, line.Quantidade, line.Escritorio)
	}

}
