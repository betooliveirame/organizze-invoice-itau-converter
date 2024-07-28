package internal

import (
	"fmt"
    "os"
    "time"
    "encoding/csv"

	"github.com/viniciusgabrielfo/organizze-invoice-itau-converter/pkg/model"
    "github.com/xuri/excelize/v2"
)

const OrganizzeSheetName = "organizze-entries.xlsx"

func GenerateOrganizzeXLXSSheet(entries []model.Entry) error {
	f := excelize.NewFile()

	defer func() error {
		if err := f.Close(); err != nil {
			return err
		}

		return nil
	}()

	_ = f.SetCellValue("Sheet1", "A1", "Data")
	_ = f.SetCellValue("Sheet1", "B1", "Descrição")
	_ = f.SetCellValue("Sheet1", "C1", "Categoria")
	_ = f.SetCellValue("Sheet1", "D1", "Valor")
	_ = f.SetCellValue("Sheet1", "E1", "Situação")

	for i := 0; i < len(entries); i++ {
		row := i + 2


		_ = f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), entries[i].Date)
		_ = f.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), entries[i].Description)
		_ = f.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), entries[i].Category)
		_ = f.SetCellFloat("Sheet1", fmt.Sprintf("D%d", row), entries[i].Value, 2, 32)
	}

	if err := f.SaveAs(OrganizzeSheetName); err != nil {
		return err
	}

	return nil
}

const OrganizzeCSVName = "organizze-entries.csv"

func GenerateOrganizzeCSV(entries []model.Entry) error {
    file, err := os.Create(OrganizzeCSVName)
    if err != nil {
        return err
    }
    defer file.Close()

    writer := csv.NewWriter(file)
    writer.Comma = ';'
    defer writer.Flush()

    // Escreve o cabeçalho no CSV
    header := []string{"Date", "Description", "Category", "Value"}
    if err := writer.Write(header); err != nil {
        return err
    }

    // Escreve as entradas no CSV
    for _, entry := range entries {
        record := entry.ToCSVRecord()
        if err := writer.Write(record); err != nil {
            return err
        }
    }

	return nil
}

const OrganizzeOFXName = "organizze-to-import.ofx"

func GenerateOrganizzeOFX(entries []model.Entry) error {

	// Abrir o arquivo para escrita
	file, err := os.Create(OrganizzeOFXName)
	if err != nil {
		return err
	}
	defer file.Close()

    // Escrever o cabeçalho OFX
	_, err = file.WriteString("OFXHEADER:100\n")
	if err != nil {
		return err
	}
	_, err = file.WriteString("DATA:OFXSGML\n")
	if err != nil {
		return err
	}
	_, err = file.WriteString("VERSION:102\n")
	if err != nil {
		return err
	}
	_, err = file.WriteString("SECURITY:NONE\n")
	if err != nil {
		return err
	}
	_, err = file.WriteString("ENCODING:USASCII\n")
	if err != nil {
		return err
	}
	_, err = file.WriteString("CHARSET:1252\n")
	if err != nil {
		return err
	}
	_, err = file.WriteString("COMPRESSION:NONE\n")
	if err != nil {
		return err
	}
	_, err = file.WriteString("OLDFILEUID:NONE\n")
	if err != nil {
		return err
	}
	_, err = file.WriteString("NEWFILEUID:NONE\n")
	if err != nil {
		return err
	}
    
    _, err = file.WriteString("<OFX>\n")
	if err != nil {
		return err
	}

    // Adicionar o bloco SIGNONMSGSRSV1
	_, err = file.WriteString("  <SIGNONMSGSRSV1>\n")
	if err != nil {
		return err
	}

	_, err = file.WriteString("    <SONRS>\n")
	if err != nil {
		return err
	}

	_, err = file.WriteString("      <STATUS>\n")
	if err != nil {
		return err
	}

	_, err = file.WriteString("        <CODE>0</CODE>\n")
	if err != nil {
		return err
	}

	_, err = file.WriteString("        <SEVERITY>INFO</SEVERITY>\n")
	if err != nil {
		return err
	}

	_, err = file.WriteString("      </STATUS>\n")
	if err != nil {
		return err
	}

// Formatar a data atual no formato YYYYMMDDHHMMSS
	currentTime := time.Now().Format("20060102")
	dtserver := "<DTSERVER>" + currentTime + "100000[-03:EST]</DTSERVER>"
	_, err = file.WriteString(fmt.Sprintf("      %s\n", dtserver))
	if err != nil {
		return err
	}
	_, err = file.WriteString("      <LANGUAGE>POR</LANGUAGE>\n")
	if err != nil {
		return err
	}

	_, err = file.WriteString("    </SONRS>\n")
	if err != nil {
		return err
	}

	_, err = file.WriteString("  </SIGNONMSGSRSV1>\n")
	if err != nil {
		return err
	}

	_, err = file.WriteString("  <BANKMSGSRSV1>\n")
	if err != nil {
		return err
	}

	_, err = file.WriteString("    <STMTTRNRS>\n")
	if err != nil {
		return err
	}

    _, err = file.WriteString("    <TRNUID>1001</TRNUID>\n")
	if err != nil {
		return err
	}
	
    _, err = file.WriteString("    <STATUS>\n")
	if err != nil {
		return err
	}

    _, err = file.WriteString("      <CODE>0</CODE>\n")
    if err != nil {
        return err
    }

    _, err = file.WriteString("      <SEVERITY>INFO</SEVERITY>\n")
    if err != nil {
        return err
    }

    _, err = file.WriteString("    </STATUS>\n")
    if err != nil {
        return err
    }

	_, err = file.WriteString("      <STMTRS>\n")
	if err != nil {
		return err
	}

    _, err = file.WriteString("        <CURDEF>BRL</CURDEF>\n")
    if err != nil {
        return err
    }

	// Adicionar informações da conta (exemplo)
	_, err = file.WriteString("        <BANKACCTFROM>\n")
	if err != nil {
		return err
	}
	_, err = file.WriteString("          <BANKID>0341</BANKID>\n")
	if err != nil {
		return err
	}
	_, err = file.WriteString("          <ACCTID>123456</ACCTID>\n")
	if err != nil {
		return err
	}
	_, err = file.WriteString("          <ACCTTYPE>CHECKING</ACCTTYPE>\n")
	if err != nil {
		return err
	}
	_, err = file.WriteString("        </BANKACCTFROM>\n")
	if err != nil {
		return err
	}

	// Adicionar transações
	_, err = file.WriteString("        <BANKTRANLIST>\n")
	if err != nil {
		return err
	}

	for _, entry := range entries {
		_, err = file.WriteString("          <STMTTRN>\n")
		if err != nil {
			return err
		}

        // Determinar o tipo de transação
		trnType := "DEBIT"
		if entry.Value > 0 {
			trnType = "CREDIT"
		}

		_, err = file.WriteString(fmt.Sprintf("            <TRNTYPE>%s</TRNTYPE>\n", trnType))
		if err != nil {
			return err
		}

		date, err := time.Parse("02/01/2006", entry.Date)
		if err != nil {
			return err
		}
		_, err = file.WriteString(fmt.Sprintf("            <DTPOSTED>%s100000[-03:EST]</DTPOSTED>\n", date.Format("20060102")))
		if err != nil {
			return err
		}

        _, err = file.WriteString(fmt.Sprintf("            <TRNAMT>%.2f</TRNAMT>\n", entry.Value))
		if err != nil {
			return err
		}

        _, err = file.WriteString(fmt.Sprintf("            <FITID>%s001</FITID>\n", date.Format("20060102")))
        if err != nil {
            return err
        }

		_, err = file.WriteString(fmt.Sprintf("            <MEMO>%s</MEMO>\n", entry.Description))
		if err != nil {
			return err
		}

        _, err = file.WriteString("            <NAME>Uber</NAME>\n")
        if err != nil {
            return err
        }

        _, err = file.WriteString("            <CATEGORY>Uber</CATEGORY>\n")
        if err != nil {
            return err
        }
		_, err = file.WriteString("          </STMTTRN>\n")
		if err != nil {
			return err
		}
	}

	_, err = file.WriteString("        </BANKTRANLIST>\n")
	if err != nil {
		return err
	}

    // Adicionar o bloco LEDGERBAL com saldo e data
	_, err = file.WriteString("        <LEDGERBAL>\n")
	if err != nil {
		return err
	}
	_, err = file.WriteString("          <BALAMT>0.00</BALAMT>\n")
	if err != nil {
		return err
	}

	currentTimeForLedger := time.Now().Format("20060102")
	_, err = file.WriteString(fmt.Sprintf("          <DTASOF>%s100000</DTASOF>\n", currentTimeForLedger))
	if err != nil {
		return err
	}
	_, err = file.WriteString(" </LEDGERBAL>\n")
	if err != nil {
		return err
	}

	_, err = file.WriteString("      </STMTRS>\n")
	if err != nil {
		return err
	}

	_, err = file.WriteString("    </STMTTRNRS>\n")
	if err != nil {
		return err
	}

	_, err = file.WriteString("  </BANKMSGSRSV1>\n")
	if err != nil {
		return err
	}


	_, err = file.WriteString("</OFX>\n")
	if err != nil {
		return err
	}

	return nil
}
