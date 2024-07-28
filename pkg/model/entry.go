package model

import (
    "fmt"
    "strings"
	//"github.com/Rhymond/go-money"
	"github.com/viniciusgabrielfo/organizze-invoice-itau-converter/pkg/category_definer"
)

type Entry struct {
	Date        string
	Description string
	Category    category_definer.Category
	Value       float64 
}

// Função para converter uma Entry para um slice de strings (para o CSV)
func (e *Entry) ToCSVRecord() []string {
    return []string{
        e.Date, // Formato padrão de data
        e.Description,
        string(e.Category),             // Ajuste conforme necessário
        strings.Replace(fmt.Sprintf("%.2f", e.Value), ".", ",", -1),
    }
}
