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

package src

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func ParseGlobalFile(file *excelize.File) map[string][]Product {
	rows, err := getRows(file, 0)

	if err != nil {
		fmt.Println("Error loading product sheet.")

		return nil
	}

	// Skip header
	rows.Next()
	headers, _ := rows.Columns()

	supplierIndex := make(map[int]string)
	result := make(map[string][]Product)

	for i := 2; i < len(headers); i++ {
		if headers[i] == "" {
			break
		}

		supplierIndex[i] = headers[i]
		result[headers[i]] = make([]Product, 0)
	}

	for rows.Next() {
		row, _ := rows.Columns()

		if len(row) < 3 || row[0] == "" || row[1] == "" {
			continue
		}

		product := &Product{
			Name:     row[0],
			Quantity: row[1],
			Price:    float64(0),
		}

		minSupplier := ""
		for i := 2; i < len(result)+2; i++ {
			price := getNumber(row[i], false)

			if price == 0 {
				continue
			}

			if product.Price == 0 || price < product.Price {
				product.Price = price
				minSupplier = supplierIndex[i]
			}
		}

		if product.Price == 0 {
			result["sem_fornecedor"] = append(result["sem_fornecedor"], *product)
		} else {
			result[minSupplier] = append(result[minSupplier], *product)
		}
	}

	return result
}

func ParseSuppliers(files []*excelize.File) []*Supplier {
	result := make([]*Supplier, 0)

	for _, file := range files {
		regex, _ := regexp.Compile("(.*[\\\\\\/])?(.+)\\.xlsx")

		name := regex.ReplaceAllString(file.Path, "$2")

		supplier := &Supplier{
			Name:                name,
			Products:            make(map[string]*Product, 0),
			TotalPricedProducts: 0,
		}

		rows, err := getRows(file, 0)

		if err != nil {
			fmt.Println("Error loading product sheet.")

			return nil
		}

		// Skip header
		rows.Next()
		headers, _ := rows.Columns()

		if len(headers) > 4 {
			continue
		}

		result = append(result, supplier)

		for rows.Next() {
			row, _ := rows.Columns()

			if len(row) == 0 {
				continue
			}

			if row[0] == "" || row[1] == "" {
				continue
			}

			var price string
			if len(row) >= 3 {
				price = row[2]
			}

			product := &Product{
				Name:     row[0],
				Quantity: row[1],
				Price:    getNumber(price, true),
			}

			supplier.Products[product.Name] = product

			if product.Price > 0 {
				supplier.TotalPricedProducts += 1
			}
		}
	}

	return result
}

func getRows(file *excelize.File, index int) (*excelize.Rows, error) {
	rows, err := file.Rows(file.GetSheetName(index))

	if err != nil || rows == nil {
		return nil, err
	}

	return rows, nil
}

func getNumber(value string, removeInvalidChars bool) float64 {
	if value == "" {
		return 0
	}

	result, err := formatNumber(value, removeInvalidChars)

	if err != nil && removeInvalidChars {
		fmt.Println("ERRO - Número em formato inválido.", err.Error())
	}

	return result
}

func formatNumber(value string, removeInvalidChars bool) (float64, error) {
	if strings.Contains(value, ",") {
		value = strings.Replace(value, ".", "", -1)
		value = strings.Replace(value, ",", ".", 1)
	}

	regex, _ := regexp.Compile("[^\\d\\.]+")

	if removeInvalidChars {
		value = regex.ReplaceAllString(value, "")
	}

	if value == "" {
		return 0, nil
	}

	return strconv.ParseFloat(value, 64)
}
