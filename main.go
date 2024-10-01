package main

import (
	"fmt"
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

	fmt.Println("What is the path of the output?")
	_, err = fmt.Scan(&destinationPath)
	err_check(err)

	// f13 from cell
	validations, err := template.GetDataValidations("Expense Report Template")
	err_check(err)
	fmt.Println(validations)

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
