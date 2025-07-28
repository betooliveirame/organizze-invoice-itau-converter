invoice_path1 ?= ./invoices/Fatura-Excel.xls
invoice_path2 ?= ./invoices/fatura.pdf
account_path ?= ./invoices/Extrato.xls

build:
	go build -v ./...

test:
	go test -race ./...

lint:
	golangci-lint run -v --fix ./...

run-invoice-itau-consumer:
	go run cmd/main.go --file $(invoice_path1) --file $(invoice_path2) --account $(account_path) --month "Agosto"

up-unoconv:
	docker buildx build -t unoconv .
	docker run -d --name unoconv unoconv

down-unoconv:
	docker container stop unoconv
	docker container rm unoconv

run-unoconv:
	docker exec unoconv unoconvert organizze-entries.xlsx organizze-entries-to-import.xls --convert-to xls
	docker cp unoconv:/organizze-entries-to-import.xls .

run: run-invoice-itau-consumer up-unoconv
	sleep 3
	docker exec unoconv unoconvert organizze-entries.xlsx organizze-entries-to-import.xls --convert-to xls
	docker cp unoconv:/organizze-entries-to-import.xls .
	$(MAKE) down-unoconv
