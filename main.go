package main

import (
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	var rosterPath, templatePath, destinationPath string
	fmt.Println("What is the path to the roster?")
	_, err := fmt.Scan(&rosterPath)
	for err != nil {
		fmt.Println(err)
		return
	}

	roster, err := excelize.OpenFile(strings.TrimSpace(rosterPath))
	if err != nil {
		fmt.Println("Error reading roster: ", err)
		return
	}
	fmt.Println(roster)

	fmt.Println("What is the path to the base report?")
	_, err = fmt.Scan(&templatePath)
	for err != nil {
		fmt.Println(err)
		return
	}

	template, err := excelize.OpenFile(strings.TrimSpace(templatePath))
	if err != nil {
		fmt.Println("Error reading template: ", err)
		return
	}
	fmt.Println(template)

	for {
		fmt.Println("What is the path of the output?")
		_, err = fmt.Scan(&destinationPath)
		if err != nil {
			fmt.Println("Error opening directory: ", err, "\n\n")
			continue
		}
		_, err = os.ReadDir(destinationPath)
		if err != nil {
			fmt.Println("Error opening directory: ", err, "\n\n")
			continue
		}
		break
	}

	locations, _ := template.GetRows("Mileage and Minutes")
	for _, row := range locations {
		fmt.Println(row)
	}
	fmt.Println("There are", len(locations), "rows")

	roster_cols, err := roster.GetCols("Sheet1")
	err_check(err)
	fmt.Println("There are", len(roster_cols), "columns in the roster")
	fmt.Println(roster_cols[2])

	today := time.Now()
	for i, name := range roster_cols[1] {
		if name == "" {
			continue
		}
		template.SetCellStr("Expense Report Template", "D7", name)
		template.SetCellStr("Expense Report Template", "D6", roster_cols[0][i])

		for j := 1; j < len(locations); j++ {
			location := get_loc_num(locations[j][0])
			roster_loc := roster_cols[2][i]
			if strings.Compare(location, roster_loc) == 0 {
				template.SetCellStr("Expense Report Template", "D8", locations[j][0])
				template.SetCellStr("Expense Report Template", "F13", locations[j][0])
				template.SetCellStr("Expense Report Template", "B13", today.Format("09/12/2024"))
				break
			}
		}

		template.SaveAs(destinationPath + "/" + name + ".xlsm")
	}
	// TODO Build ui
}

func err_check(err error) {
	for err != nil {
		fmt.Println(err)
		return
	}
}

func get_loc_num(location string) string {
	var result string
	for _, character := range location {
		if character == '#' {
			continue
		}
		if character == ' ' {
			break
		}
		result += string(character)
	}

	return result
}
