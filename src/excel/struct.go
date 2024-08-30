package excel

import "time"

// Struct contendo os campos da tabela
// ATENÇÂO
// O xlsx deve conter exatamente o nome da coluna que coresponde a o campo da struct
// EX:
//`xlsx:"AQUI VC COLOCA O NOME DA COLUNA"`
type linha struct {
	Conta      int       `xlsx:"NUMERO DA CONTA"`
	Nome       string    `xlsx:"NOME"`
	Quantidade int       `xlsx:"QUANTIDADE"`
	Data       time.Time `xlsx:"DATA"`
	Telefone   string    `xlsx:"CEL"`
}
