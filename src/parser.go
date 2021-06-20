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

func ParseSuppliers(files []*excelize.File) []*Supplier {
	result := make([]*Supplier, 0)

	for _, file := range files {
		regex, _ := regexp.Compile(".+[\\\\\\/](.+?).xlsx")

		name := regex.ReplaceAllString(file.Path, "$1")

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

			if len(row) < 3 || row[0] == "" || row[1] == "" {
				continue
			}

			product := &Product{
				Name:     row[0],
				Quantity: row[1],
				Price:    getNumber(row[2]),
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

func isNumber(value string) bool {
	_, err := formatNumber(value)

	return err == nil
}

func getNumber(value string) float64 {
	result, err := formatNumber(value)

	if err != nil {
		fmt.Println("Error parsing value.", err.Error())
	}

	return result
}

func formatNumber(value string) (float64, error) {
	if strings.Contains(value, ",") {
		value = strings.Replace(value, ".", "", -1)
		value = strings.Replace(value, ",", ".", 1)
	}

	value = strings.Replace(value, "%", "", 1)

	regex, _ := regexp.Compile("[^\\d\\.]+")

	value = regex.ReplaceAllString(value, "")

	if value == "" {
		return 0, nil
	}

	return strconv.ParseFloat(value, 64)
}
