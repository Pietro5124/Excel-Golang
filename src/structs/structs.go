package structs

import "time"

// Struct contendo os campos da tabela
// ATENÇÂO
// O xlsx deve conter exatamente o nome da coluna que coresponde a o campo da struct
// EX:
//`xlsx:"AQUI VC COLOCA O NOME DA COLUNA"`
type Nnm struct {
	Conta      int       `xlsx:"'DL_F_MovimentaçãoCosolidada'[COD_CARTEIRA]"`
	Nome       string    `xlsx:"NOME"`
	Quantidade int       `xlsx:"QUANTIDADE"`
	Data       time.Time `xlsx:"DATA"`
	Escritorio string    `xlsx:"ESCRITÓRIO"`
}
