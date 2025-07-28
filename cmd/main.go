package main

import (
	"flag"
	"fmt"
	"log/slog"
	"time"
	"sort"
	"strings"
	"path/filepath"

	"invoice-processor/internal"
	"invoice-processor/pkg/model"	
)

// StringSlice é um tipo customizado para aceitar múltiplos valores do mesmo flag
type StringSlice []string

func (s *StringSlice) String() string {
	return strings.Join(*s, ",")
}

func (s *StringSlice) Set(value string) error {
	*s = append(*s, value)
	return nil
}

var (
	invoicePaths StringSlice
	accountPath  string
	month        string
	startDate    string
	endDate      string
)

func init() {
	flag.Var(&invoicePaths, "file", "itaú invoice path to consume (can be used multiple times)")
	flag.StringVar(&accountPath, "account", "", "itaú account path to consume")
	flag.StringVar(&month, "month", "", "month to generate the file (02/2006)")
	flag.StringVar(&startDate, "start-date", "", "only consume from start-date (02/01/2006)")
	flag.StringVar(&endDate, "end-date", "", "only consume until end-date (02/01/2006)")
	flag.Parse()
}

// isPDFFile verifica se o arquivo é PDF pela extensão
func isPDFFile(filePath string) bool {
	ext := strings.ToLower(filepath.Ext(filePath))
	return ext == ".pdf"
}

func main() {
	l := slog.Default()
	l.Info("Starting invoice-itau-consumer...")

	// Validar se pelo menos um arquivo de invoice foi fornecido
	if len(invoicePaths) == 0 {
		l.Error("At least one invoice file must be provided using -file flag")
		return
	}

	itauImportConfigs := &internal.ItauImportConfigs{}

	if startDate != "" {
		tStartDate, err := time.Parse("02/01/2006", startDate)
		if err != nil {
			panic(err)
		}

		itauImportConfigs.StartDate = tStartDate
	}

	if endDate != "" {
		tEndDate, err := time.Parse("02/01/2006", endDate)
		if err != nil {
			panic(err)
		}

		itauImportConfigs.EndDate = tEndDate
	}

	entries := make([]model.Entry, 0)
	var err error
	
	// Processar múltiplos arquivos de invoice
	for _, invoicePath := range invoicePaths {
		l.Info(fmt.Sprintf("Processing invoice: %s", invoicePath))
		
		// Detectar automaticamente se é PDF pela extensão
		if isPDFFile(invoicePath) {
			l.Info(fmt.Sprintf("Detected PDF file: %s", invoicePath))
			entries, err = internal.GetEntriesFromItauInvoiceFromPDF(entries, itauImportConfigs, invoicePath)
		} else {
			l.Info(fmt.Sprintf("Detected XLS file: %s", invoicePath))
			entries, err = internal.GetEntriesFromItauInvoice(entries, itauImportConfigs, invoicePath)
		}
		
		if err != nil {
			l.Error(fmt.Sprintf("Error processing %s: %s", invoicePath, err.Error()))
			return
		}
		
		l.Info(fmt.Sprintf("Successfully processed %s with %d entries", invoicePath, len(entries)))
	}

	entries, err = internal.GetEntriesFromItauAccount(entries, itauImportConfigs, accountPath)
	if err != nil {
		l.Error(err.Error())
		return
	}

	l.Info(fmt.Sprintf("All %d invoice files successfully processed!", len(invoicePaths)))
	l.Info(fmt.Sprintf("Starting to generate %s...", internal.OrganizzeOFXName))

	// Ordena as entradas pelo campo date
	sort.Slice(entries, func(i, j int) bool {
		return entries[i].Date < entries[j].Date
	})

	if err := internal.GenerateOrganizzeXLXSSheet(month, entries); err != nil {
		l.Error(err.Error())
		return
	}

	l.Info(fmt.Sprintf("Finished, %s was generated!", internal.OrganizzeSheetName))
}
