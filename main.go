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

	fmt.Println("What is the path of the output?")
	_, err = fmt.Scan(&destinationPath)
	for err != nil {
		fmt.Println(err)
		return
	}

	/* Next steps:
	Create array of names, parse names into first and last
	Create array of locations and EE#
	Create a copy of expense reports in output directory
	Rename and fill out the expense reports
	*/
}
