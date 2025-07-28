package category_definer

import "strings"

type Category string

var (
	Alimentacao = Category("Alimentação")
	AssinaturasEServicosDeMusica = Category("Assinaturas e Serviços - Música")
	AssinaturasEServicosDeVideo = Category("Assinaturas e Serviços - Video")
	AssinaturasEServicosOutros = Category("Assinaturas e Serviços - Outros")
	BaresERestaurantes = Category("Bares e Restaurantes")
	Casa = Category("Casa")
	CasaFinanciamentoImobiliario = Category("Casa - Financiamento Imobiliário")
	CasaInternet = Category("Casa - Internet")
	CasaAgua = Category("Casa - Conta de água")
	CasaEnergia = Category("Casa - Conta de energia")
	CasaManutencao = Category("Casa - Manutenção da casa")
	CasaGas = Category("Casa - Conta de gás")
	CasaSeguroResidencial = Category("Casa - Seguro Residencial")
	Compras = Category("Compras")
	CuidadosPessoais = Category("Cuidados Pessoais")
	Educacao = Category("Educacao")
	FamiliaEFilhos = Category("Familia e Filhos")
	ImpostosETaxas = Category("Impostos e Taxas")
	Investimentos = Category("Investimentos")
	LazerEHobbies = Category("Lazer e Hobbies")
	Mercado = Category("Mercado")
	MercadoShopper = Category("Mercado - Shopper")
	MercadoFeira = Category("Mercado - Feira")
	PresentesEDoacoes = Category("Presentes e Doações")
	Roupas = Category("Roupas")
	Saude = Category("Saúde")
	Trabalho = Category("Trabalho")
	Transporte = Category("Transporte")
	TransporteUber = Category("Transporte - Uber")
	Viagem = Category("Viagem")
	Salario = Category("Salário")
	Reembolso = Category("Reembolso")
	FaturaDoCartaoDeCredito = Category("Fatura do cartão de crédito")
)

var extenseCategoryKeyWords = map[Category][]string{
	Alimentacao: {"Mini Extra", "Ifd*", "Bon Forno", "Grao Espresso", "CASA DE BOLO", "NOVA ITALIANA", "MANIA DE CHURRASCO", "SHOPPING ABC", "PADARIA PARATI"},
	AssinaturasEServicosDeMusica: {"Spotify"},
	AssinaturasEServicosDeVideo: {"Netflix", "Amazon Prime", "YOUTUB"},
	AssinaturasEServicosOutros: {"Apple", "Mp *melimais", "NIRVANAHQ", "X CORP", "F.G.G. PELINSON SERVI"},
	BaresERestaurantes: {"BAR", "Beco do espeto","Rancho","ADEGA", "Tios", "Pit Stop Aurea", "Restaurante", "Quiosque", "PIT STOP ADRIATICO", "BUBBLE MIX ABC", "ZE DELIVERY", "MARQUINHOS BAR", "RECANTO O CASULO", "CWBarLtda"},
	Compras: {"Fb Servicos", "Mp *", "Bazar Casa Da Mamae", "Multicoisas", "Mercado Mateus", "Mercadopago", "DAISO BRASIL", "MERCADOLIVRE"},
	Mercado: {"Rosetti", "Nagumo", "Rossetti", "COOP"},
	MercadoShopper: {"Shopper"},
	MercadoFeira: {"Feira"},
	Roupas: {"Torra", "Besni"},
	LazerEHobbies: {"sony"},
	ImpostosETaxas: {"Anuidade Diferenc", "JUROS LIMITE DA CONTA", "SEGURO CARTAO"},
	CasaManutencao: {"Pereiramix", "Pereira Mix", "Mix Com Ferragens"},
	CasaSeguroResidencial: {"SISDEB  ITAUPORTOSEGUR"},
	CasaInternet: {"CLARO"},
	CasaFinanciamentoImobiliario: {"FINANC IMOBILIARIO"},
	TransporteUber: {"Uber"},
	Viagem: {"Hotel", "Smiles Clube Smiles"},
	Saude: {"The King Gym", "DROGARIA"},
	FaturaDoCartaoDeCredito: {"ITAU MC", "ITAU BLACK"},
	FamiliaEFilhos: {"PIX TRANSF  HELISON"},
	CuidadosPessoais: {"MEGA COLOR"},
}

var incomeCategoryKeyWords = map[Category][]string{
	Salario: {"SALARIO"},
	Reembolso: {"PIX"},
}

func GetCategoryFromDescription(description string, value float64) Category {
	if value > 0 {
		return GetCategoryFromDescriptionIncome(description)
	}

	return GetCategoryFromDescriptionExpense(description)
}

func GetCategoryFromDescriptionExpense(description string) Category {
	for category, keys := range extenseCategoryKeyWords {
		for i := 0; i < len(keys); i++ {
			if strings.Contains(strings.ToLower(description), strings.ToLower(keys[i])) {
				return category
			}
		}
	}

	return ""
}

func GetCategoryFromDescriptionIncome(description string) Category {
	for category, keys := range incomeCategoryKeyWords {
		for i := 0; i < len(keys); i++ {
			if strings.Contains(strings.ToLower(description), strings.ToLower(keys[i])) {
				return category
			}
		}
	}

	return ""
}
