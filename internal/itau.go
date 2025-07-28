package internal

import (
	"errors"
	"log/slog"
	"regexp"
	"strconv"
	"strings"
	"time"
	"slices"
	"fmt"

	"invoice-processor/pkg/category_definer"
	"invoice-processor/pkg/model"
	"github.com/viniciusgabrielfo/xls"
	"github.com/ledongthuc/pdf"
)

type ItauImportConfigs struct {
	StartDate time.Time
	EndDate   time.Time
}

func GetEntriesFromItauInvoice(entries []model.Entry, configs *ItauImportConfigs, filePath string) ([]model.Entry, error) {
	logger := slog.Default()

	f, err := xls.Open(filePath, "utf-8")
	if err != nil {
		return nil, err
	}

	sheet := f.GetSheet(0)

	if sheet == nil {
		return nil, errors.New("invalid sheet")
	}

	var isEntry bool

	for i := 0; i <= int(sheet.MaxRow); i++ {
		row := sheet.Row(i)
		if row == nil {
			if isEntry {
				isEntry = false
			}
			continue
		}

		date := row.Col(0)
		description := row.Col(1)

		if date == "data" && description == "lançamento" {
			isEntry = true
			continue
		}

		if isEntry {
			if date == "" || description == "dólar de conversão" {
				continue
			}

			entryDate, err := time.Parse("02/01/2006", date)
			if err != nil {
				logger.Error(err.Error())
				continue
			}

			if !IsBetweenConfigInternal(configs, entryDate) {
				continue
			}

			value, err := strconv.ParseFloat(row.Col(3), 64)
			value = -value
			if err != nil {
				return entries, err
			}

			// if ok, installments := IsInstallmentPurchase(description); ok {
			// 	value = value * float64(installments)
			// }

			entries = append(entries, model.Entry{
				Date:        date,
				Description: description,
				Category:    category_definer.GetCategoryFromDescriptionExpense(description),
				Value:       value,
				Type:        "cartão de crédito",
			})
		}
	}

	return entries, nil
}

func GetEntriesFromItauAccount(entries []model.Entry, configs *ItauImportConfigs, filePath string) ([]model.Entry, error) {
	logger := slog.Default()

	descritpionsToSkip := []string{"SALDO DO DIA"}
	dateToSkip := []string{"lançamentos", "lançamentos futuros", "saídas futuras"}

	f, err := xls.Open(filePath, "utf-8")
	if err != nil {
		return nil, err
	}

	sheet := f.GetSheet(0)

	if sheet == nil {
		return nil, errors.New("invalid sheet")
	}

	var isEntry bool

	for i := 0; i <= int(sheet.MaxRow); i++ {
		row := sheet.Row(i)
		if row == nil {
			if isEntry {
				isEntry = false
			}
			continue
		}

		date := row.Col(0)
		description := row.Col(1)
		
		if date == "data" && description == "lançamento" {
			isEntry = true
			continue
		}
		
		if isEntry {
			if slices.Contains(dateToSkip, date) {
				continue
			}

			if date == "" || description == "dólar de conversão" {
				continue
			}

			if slices.Contains(descritpionsToSkip, description) {
				continue
			}

			entryDate, err := time.Parse("02/01/2006", date)
			if err != nil {
				logger.Error(err.Error())
				continue
			}

			if !IsBetweenConfigInternal(configs, entryDate) {
				continue
			}

			value, err := strconv.ParseFloat(row.Col(3), 64)
			if err != nil {
				logger.Error(row.Col(3))
				return entries, err
			}

			// if ok, installments := IsInstallmentPurchase(description); ok {
			// 	value = value * float64(installments)
			// }

			entries = append(entries, model.Entry{
				Date:        date,
				Description: description,
				Category:    category_definer.GetCategoryFromDescription(description, value),
				Value:       value,
				Type:        "conta corrente",
			})
		}
	}

	return entries, nil
}

func GetEntriesFromItauInvoiceFromPDF(entries []model.Entry, configs *ItauImportConfigs, filePath string) ([]model.Entry, error) {
	logger := slog.Default()

	f, r, err := pdf.Open(filePath)
	if err != nil {
		logger.Error(err.Error())
		return entries, err	
	}
	defer f.Close()
	
	var fullText strings.Builder
	
	// Extrair texto de todas as páginas
	for pageNum := 1; pageNum <= r.NumPage(); pageNum++ {
		p := r.Page(pageNum)
		if p.V.IsNull() {
			continue
		}
		
		text, err := p.GetPlainText(nil)
		if err != nil {
			continue
		}
		fullText.WriteString(text)
		fullText.WriteString("\n")
	}

	text := fullText.String()

	// Excluir o que estiver acima de "Lançamentos: compras e saques"
	start := strings.Index(text, "Lançamentos: compras e saques")
	if start == -1 {
		return entries, err
	}
	text = text[start:]

	end := strings.Index(text, "Compras parceladas - próximas faturas")
	if end == -1 {
		return entries, err
	}
	text = text[:end]
	
	// Extrair transações
	re := regexp.MustCompile(`(\d{2}\/\d{2})\s*\n\s*([^\n]+)\s*\n\s*(-?\s*[\d,]+)`)
	matches := re.FindAllStringSubmatch(text, -1)
	
	for _, match := range matches {
		if len(match) >= 4 {
			date := match[1]
			description := strings.TrimSpace(match[2])
			amountStr := strings.ReplaceAll(match[3], " - ", "-")
			amountStr = strings.ReplaceAll(amountStr, " -", "-")
			amountStr = strings.ReplaceAll(amountStr, "- ", "-")
			
			// Limpar e converter valor
			amountStr = strings.ReplaceAll(amountStr, ".", "")
			amountStr = strings.ReplaceAll(amountStr, ",", ".")
			value, err := strconv.ParseFloat(amountStr, 64)
			value = -value
			if err != nil {
				logger.Error(err.Error())
				return entries, err
			}
			
			// Adicionar ano atual se não estiver presente
			if !strings.Contains(date, "/2") {
				currentYear := time.Now().Year()
				date = fmt.Sprintf("%s/%d", date, currentYear)
			}

			entries = append(entries, model.Entry{
				Date:        date,
				Description: description,
				Category:    category_definer.GetCategoryFromDescriptionExpense(description),
				Value:       value,
				Type:        "cartão de crédito",
			})
		}
	}
	
	return entries, nil	
}

func IsBetweenConfigInternal(configs *ItauImportConfigs, date time.Time) bool {
	if !configs.StartDate.IsZero() {
		if date.Before(configs.StartDate) {
			return false
		}
	}

	if !configs.EndDate.IsZero() {
		if date.After(configs.EndDate) {
			return false
		}
	}

	return true
}

func IsInstallmentPurchase(description string) (bool, int32) {
	logger := slog.Default()

	re, err := regexp.Compile("01/[0-9]+")
	if err != nil {
		logger.Error(err.Error())
		return false, 0
	}

	installmentPattern := re.FindAllString(description, -1)
	if len(installmentPattern) == 0 {
		return false, 0
	}

	s := strings.Split(installmentPattern[0], "/")

	i, err := strconv.ParseInt(s[1], 10, 32)
	if err != nil {
		logger.Error(err.Error())
		return false, 0
	}

	return true, int32(i)
}
