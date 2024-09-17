package main

import "fmt"
import "github.com/xuri/excelize/v2"

func main() {
	var rosterPath, templatePath, destinationPath string
	fmt.Println("What is the path to the roster?")
	_, err := fmt.Scan(&rosterPath)
	for err != nil {

	}
	fmt.Println("What is the path to the base report?")
	fmt.Scan(&templatePath)
	fmt.Println("What is the path of the output?")
	fmt.Scan(&destinationPath)

	
}
