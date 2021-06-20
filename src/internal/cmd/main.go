// Copyright 2021 Lucas Soares
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

package main

import (
	"bufio"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/jessevdk/go-flags"
	"github.com/lucasoares/simple-supplier-processor/src"
)

type Opts struct {
	Folder string `short:"f" long:"folder" description:"folder to process" required:"false" value-name:"FOLDER"`
}

func main() {
	var opts Opts

	_, err := flags.ParseArgs(&opts, os.Args[1:])

	if err != nil {
		exitWithError(err)
	}

	if opts.Folder == "" {
		opts.Folder = "./"
	}

	if !strings.HasSuffix(opts.Folder, "/") {
		fmt.Println("ERRO - O argumento -f precisa possuir um diretório finalizando em '/'.")
		exitProgram()
	}

	filesCount := getDirFilesCount("resultado/")

	if filesCount == 0 {
		fmt.Println("Computando arquivo geral.")
		fmt.Println()

		os.Mkdir("resultado/", 0755)

		computeGlobalSheet(opts.Folder)
	} else if filesCount == 1 {
		fmt.Println("Computando resultado de fornecedores.")

		computeResultSheets(opts.Folder)
	} else {
		fmt.Println("ERRO - Diretório 'resultado' ja está totalmente preenchido.\nExclua os arquivos com exceção do arquivo 'resultado/resultado_geral.xlsx' e rode novamente.")
	}

	exitProgram()
}

func computeResultSheets(folder string) {
	file, err := excelize.OpenFile("resultado/resultado_geral.xlsx")

	if err != nil {
		exitWithError(err)
	}

	result := src.ParseGlobalFile(file)

	totalSuppliers := len(result)
	if _, ok := result["sem_fornecedor"]; ok {
		totalSuppliers--
	}

	fmt.Println()
	if totalSuppliers == 1 {
		fmt.Println("1 fornecedor no resultado:")
	} else {
		fmt.Println(fmt.Sprintf("%d fornecedores no resultado:", totalSuppliers))
	}
	fmt.Println()

	for i := range result {
		if i == "sem_fornecedor" {
			continue
		}

		fmt.Println(fmt.Sprintf("%s", i))
		fmt.Println(fmt.Sprintf("%d produtos", len(result[i])))
		fmt.Println()
	}

	if data, ok := result["sem_fornecedor"]; ok {
		fmt.Println(fmt.Sprintf("Produtos sem fornecedor disponível: %d.", len(data)))
	}
	fmt.Println()

	writeFinalResult(result)
}

func writeFinalResult(result map[string][]src.Product) {
	for n, products := range result {
		f := excelize.NewFile()
		decimalStyle, _ := f.NewStyle(&excelize.Style{DecimalPlaces: 2, NumFmt: 4})

		f.SetSheetName("Sheet1", n)
		f.SetActiveSheet(0)

		f.SetCellValue(n, "A1", "PRODUTO")
		f.SetCellValue(n, "B1", "QTD")
		f.SetCellValue(n, "C1", "PREÇO")

		f.SetColWidth(n, "A", "A", 60)
		f.SetColWidth(n, "B", "B", 8)
		f.SetColWidth(n, "C", "C", 14)
		f.SetRowHeight(n, 1, 30)

		f.SetColStyle(n, "C", decimalStyle)

		for i, product := range products {
			f.SetCellValue(n, fmt.Sprintf("A%d", i+2), product.Name)
			f.SetCellValue(n, fmt.Sprintf("B%d", i+2), product.Quantity)
			f.SetCellValue(n, fmt.Sprintf("C%d", i+2), product.Price)
		}

		// Save
		fileName := fmt.Sprintf("resultado/%s.xlsx", n)
		if err := f.SaveAs(fileName); err != nil {
			fmt.Println("ERRRO - Erro ao salvar arquivo de resultado. Verifique suas permissões.")

			exitProgram()
		}

		fmt.Println(fmt.Sprintf("Arquivo '%s' foi criado com sucesso.", fileName))
	}
}

func computeGlobalSheet(folder string) {
	paths, err := filepath.Glob(folder + "*.xlsx")

	if err != nil {
		exitWithError(err)
	}

	fmt.Println(fmt.Sprintf("%d arquivos encontrados:", len(paths)))
	for i := range paths {
		fmt.Println(fmt.Sprintf("%v", paths[i]))
	}
	fmt.Println()

	files := make([]*excelize.File, 0)
	for _, path := range paths {
		file, err := excelize.OpenFile(path)

		if err != nil {
			exitWithError(err)

			return
		}

		files = append(files, file)
	}

	suppliers := src.ParseSuppliers(files)

	if len(suppliers) == 1 {
		fmt.Println("1 fornecedor encontrado:")
	} else {
		fmt.Println(fmt.Sprintf("%d fornecedores encontrados:", len(suppliers)))
	}

	for i := range suppliers {
		fmt.Println(fmt.Sprintf("%s", suppliers[i].Name))
	}
	fmt.Println()

	fmt.Println("Resumo dos fornecedores:")
	fmt.Println()

	lenProducts := len(suppliers[0].Products)
	for _, supplier := range suppliers {
		fmt.Println(fmt.Sprintf("%v", supplier.Name))
		fmt.Println(fmt.Sprintf("%d produtos", len(supplier.Products)))
		fmt.Println(fmt.Sprintf("%d produtos precificados", supplier.TotalPricedProducts))

		if len(supplier.Products) != lenProducts {
			fmt.Println(fmt.Sprintf("ATENÇÃO - Número de produtos diferente de outras planilhas: %d", len(supplier.Products)))
		}
		fmt.Println()
	}

	if len(suppliers) > 15 {
		fmt.Println("ERRO - O sistema suporta no momento apenas um máximo de 15 fornecedores.")
		exitProgram()
	}

	writeGlobalSheet(suppliers)
}

func writeGlobalSheet(suppliers []*src.Supplier) {
	f := excelize.NewFile()
	alignmentStyle, _ := f.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{JustifyLastLine: true, ReadingOrder: 0, RelativeIndent: 1, ShrinkToFit: true, WrapText: true}})
	decimalStyle, _ := f.NewStyle(&excelize.Style{DecimalPlaces: 2, NumFmt: 4})

	redStyle, _ := f.NewConditionalStyle(`{"fill":{"type":"pattern","color":["ff0000"],"pattern":1}}`)
	yellowStyle, _ := f.NewConditionalStyle(`{"fill":{"type":"pattern","color":["ffff00"],"pattern":1}}`)
	greenStyle, _ := f.NewConditionalStyle(`{"fill":{"type":"pattern","color":["#00FF00"],"pattern":1}}`)

	n := "Geral"

	f.SetSheetName("Sheet1", n)
	f.SetActiveSheet(0)

	f.SetCellValue(n, "A1", "PRODUTO")
	f.SetCellValue(n, "B1", "QTD")

	f.SetColWidth(n, "A", "A", 60)
	f.SetColWidth(n, "B", "B", 8)
	f.SetColWidth(n, "C", "Z", 14)
	f.SetRowHeight(n, 1, 30)

	f.SetColStyle(n, "B:Z", alignmentStyle)

	// Write product names
	products := make([]string, len(suppliers[0].Products))
	i := 0
	for k := range suppliers[0].Products {
		products[i] = k
		i++
	}

	sort.Strings(products)

	for i := range products {
		f.SetCellValue(n, fmt.Sprintf("A%d", i+2), products[i])
		f.SetCellValue(n, fmt.Sprintf("B%d", i+2), suppliers[0].Products[products[i]].Quantity)
	}

	// Write supplier prices
	supplierIndex := 0
	var supplierColumn rune
loop:
	for supplierColumn = 'C'; supplierColumn <= 'Q'; supplierColumn++ {
		if supplierIndex >= len(suppliers) {
			break loop
		}

		supplier := suppliers[supplierIndex]

		f.SetCellValue(n, fmt.Sprintf("%c1", supplierColumn), supplier.Name)

		for i := range products {
			price := supplier.Products[products[i]].Price

			if price > 0 {
				f.SetCellValue(n, fmt.Sprintf("%c%d", supplierColumn, i+2), price)
			}
		}

		supplierIndex++
	}

	supplierColumn--

	// Best price
	bestPriceColumn := supplierColumn + 3
	f.SetCellValue(n, fmt.Sprintf("%c1", bestPriceColumn), "Melhor Preço")

	for i := 2; i < len(products)+2; i++ {
		f.SetCellFormula(n, fmt.Sprintf("%c%d", bestPriceColumn, i), fmt.Sprintf("MIN(C%d:%c%d)", i, supplierColumn, i))
	}

	// Best supplier
	bestSupplierColumn := supplierColumn + 4
	f.SetCellValue(n, fmt.Sprintf("%c1", bestSupplierColumn), "Fornecedor")

	for i := 2; i < len(products)+2; i++ {
		f.SetCellFormula(n, fmt.Sprintf("%c%d", bestSupplierColumn, i), fmt.Sprintf("INDEX($C$1:$%c$1,MATCH(MIN(C%d:%c%d),C%d:%c%d,0))", supplierColumn, i, supplierColumn, i, i, supplierColumn, i))
	}

	// Worse price
	worsePriceColumn := supplierColumn + 5
	f.SetCellValue(n, fmt.Sprintf("%c1", worsePriceColumn), "Pior Preço")

	for i := 2; i < len(products)+2; i++ {
		f.SetCellFormula(n, fmt.Sprintf("%c%d", worsePriceColumn, i), fmt.Sprintf("MAX(C%d:%c%d)", i, supplierColumn, i))
	}

	// Worse minus best price
	diffPriceColumn := supplierColumn + 6
	f.SetCellValue(n, fmt.Sprintf("%c1", diffPriceColumn), "Diferença")

	for i := 2; i < len(products)+2; i++ {
		cell := fmt.Sprintf("%c%d", diffPriceColumn, i)
		f.SetCellFormula(n, cell, fmt.Sprintf("MAX(C%d:%c%d)-MIN(C%d:%c%d)", i, supplierColumn, i, i, supplierColumn, i))
		f.SetCellStyle(n, cell, cell, decimalStyle)
	}

	// Percentage of price difference
	pricePercentageColumn := supplierColumn + 7
	f.SetCellValue(n, fmt.Sprintf("%c1", pricePercentageColumn), "% Diferença")

	for i := 2; i < len(products)+2; i++ {
		cell := fmt.Sprintf("%c%d", pricePercentageColumn, i)
		f.SetCellFormula(n, cell, fmt.Sprintf("100*%c%d/%c%d", diffPriceColumn, i, bestPriceColumn, i))
		f.SetCellStyle(n, cell, cell, decimalStyle)
	}
	f.SetConditionalFormat(n, fmt.Sprintf("%c2:%c%d", pricePercentageColumn, pricePercentageColumn, len(products)+1), fmt.Sprintf(`[{"type":"cell","criteria":">","format":%d,"value":"50"}]`, redStyle))
	f.SetConditionalFormat(n, fmt.Sprintf("%c2:%c%d", pricePercentageColumn, pricePercentageColumn, len(products)+1), fmt.Sprintf(`[{"type":"cell","criteria":">","format":%d,"value":"20"}]`, yellowStyle))
	f.SetConditionalFormat(n, fmt.Sprintf("%c2:%c%d", pricePercentageColumn, pricePercentageColumn, len(products)+1), fmt.Sprintf(`[{"type":"cell","criteria":"<=","format":%d,"value":"20"}]`, greenStyle))

	// Auto Filter
	f.AutoFilter(n, "A1", fmt.Sprintf("%c1", pricePercentageColumn), "")

	// Save
	os.Mkdir("resultado", 0755)

	if err := f.SaveAs("resultado/resultado_geral.xlsx"); err != nil {
		fmt.Println("ERRRO - Erro ao salvar arquivo de resultado. Verifique suas permissões.")

		exitProgram()
	}

	fmt.Println("Arquivo 'resultado/resultado_geral.xlsx' foi criado com sucesso.")
}

func exitWithError(err error) {
	fmt.Println(fmt.Sprintf("%v", err.Error()))
	fmt.Println()
	fmt.Println("ERRO - Erro ao executar a aplicação.")
	exitProgram()
}

func exitProgram() {
	fmt.Println()
	fmt.Print("Aperte 'Enter' para finalizar...")
	bufio.NewReader(os.Stdin).ReadBytes('\n')

	os.Exit(0)
}

func isEmpty(name string) bool {
	f, err := os.Open(name)
	if err != nil {
		return true
	}
	defer f.Close()

	_, err = f.Readdirnames(1)
	if err == io.EOF {
		return true
	}
	return false
}

func getDirFilesCount(name string) int {
	f, err := os.Open(name)
	if err != nil {
		return 0
	}
	defer f.Close()

	names, err := f.Readdirnames(-1)
	if err == io.EOF {
		return 0
	}
	return len(names)
}
