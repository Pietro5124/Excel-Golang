package excel

import (
	"fmt"
	"io"
	"main/src/structs"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func OpenExcel(file io.Reader) ([]structs.Linha, error) {
	// Linha de Inicio que o excel começa a ler obs: A lista inicia a contagem de linha pelo 0
	var StarLine int = 2

	// Abrir o arquivo Excel
	excel, err := excelize.OpenReader(file)
	if err != nil {
		// Tratar o erro
		return []structs.Linha{}, err
	}

	var lines []structs.Linha

	rows, err := excel.GetRows("Sheet1")
	if err != nil {
		return []structs.Linha{}, err
	}
	// Aqui vc define apartir de que linha começa a tabela
	headerRow := rows[StarLine]

	fieldMap := getFieldMap(reflect.TypeOf(structs.Linha{}))

	// Encontra o índice da coluna para cada campo da struct pelo nome
	fieldIndexes := make(map[string]int)
	for colIndex, colName := range headerRow {
		for fieldName, tagValue := range fieldMap {
			// Compara o nome contido na coluna sem espaços com a tags xlsx
			if strings.EqualFold(strings.ReplaceAll(tagValue, " ", ""), strings.ReplaceAll(colName, " ", "")) {
				fieldIndexes[fieldName] = colIndex
				break
			}
		}
	}

	// Lê os dados da planilha iniciando pela linha de inicio e preenche a struct
	for _, row := range rows[StarLine+1:] {
		line := structs.Linha{}

		for fieldName, fieldIndex := range fieldIndexes {
			// Pega o campo da Struct pelo nome da mesma
			fieldValue := reflect.ValueOf(&line).Elem().FieldByName(fieldName)
			// Verifica se o compo existe
			if fieldValue.IsValid() && fieldValue.CanSet() {

				// Verifica o tipo do mesmo
				switch fieldValue.Kind() {
				case reflect.String:

					fieldValue.SetString(row[fieldIndex])

				case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
					idade, err := parseInt(row[fieldIndex])
					if err == nil {
						fieldValue.SetInt(int64(idade))
					}
				case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
					valor, err := parseValor(row[fieldIndex])
					if err == nil {
						fieldValue.SetUint(uint64(valor))
					}
				case reflect.Float32, reflect.Float64:
					valor, err := parseValor(row[fieldIndex])
					if err == nil {
						fieldValue.SetFloat(float64(valor))
					}
				case reflect.Bool:
					valor, err := parseBool(row[fieldIndex])
					if err == nil {
						fieldValue.SetBool(valor)
					}
				default:

					if fieldValue.Type() == reflect.TypeOf(time.Time{}) {

						data, err := parseTime(row[fieldIndex])
						if err == nil {

							fieldValue.Set(reflect.ValueOf(data))
						}
					}

				}
			}
		}

		lines = append(lines, line)
	}
	return lines, nil
}

func getFieldMap(structType reflect.Type) map[string]string {
	fieldMap := make(map[string]string)
	for i := 0; i < structType.NumField(); i++ {
		field := structType.Field(i)
		tagValue := field.Tag.Get("xlsx")
		if tagValue != "" {
			fieldMap[field.Name] = tagValue
		}
	}
	return fieldMap
}

func parseTime(data string) (time.Time, error) {

	// Caso a conversão com o formato "1/27/2006"
	parsedTime, err := time.Parse("02/01/2006", data)
	if err == nil {
		return parsedTime, nil
	}
	// Tente fazer a conversão com o formato "2006-01-02"
	parsedTime, err = time.Parse("2006-01-02", data)
	if err == nil {
		return parsedTime, nil
	}

	// Se nenhum formato funcionar, retorne um erro
	return time.Time{}, fmt.Errorf("Formato de data inválido: %s", data)
}

func parseInt(s string) (int64, error) {
	valor, err := strconv.ParseInt(s, 10, 64)
	if err != nil {
		return 0, err
	}
	return valor, nil
}

func parseValor(s string) (float64, error) {
	valor, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return 0, err
	}
	return valor, nil
}

func parseBool(s string) (bool, error) {
	valor, err := strconv.ParseBool(s)
	if err != nil {
		return false, err
	}
	return valor, nil
}
