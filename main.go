package main

import (
	"fmt"
	"os"
	"strings"

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

	template, err := excelize.OpenFile(strings.TrimSpace(rosterPath))
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

	firstSheet := roster.GetSheetList()[0]
	roster_cols, err := roster.GetCols(firstSheet)
	err_check(err)

	for i, name := range roster_cols[1] {
		if name == "" {
			continue
		}
		template.SetCellStr("Expense Report Template", "D7", name)
		template.SetCellStr("Expense Report Template", "D6", roster_cols[0][i])

		for j := 1; j < len(locations); j++ {
			if locations[j][0][1:2] == roster_cols[2][i+3] {
				template.SetCellStr("Expense Report Template", "D8", locations[0][j])
				template.SetCellStr("Expense Report Template", "F13", locations[0][j])
			}
		}

		template.SaveAs(destinationPath + "/" + name + ".xlsm")
	}

	/* Next steps:
	Create array of names, parse names into first and last
	Create array of locations and EE#
	Create a copy of expense reports in output directory
	Rename and fill out the expense reports
	*/
}

func err_check(err error) {
	for err != nil {
		fmt.Println(err)
		return
	}
}
