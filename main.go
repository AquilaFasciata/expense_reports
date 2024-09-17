package main

import "fmt"
import "github.com/xuri/excelize/v2"

func main() {
	var rosterPath, templatePath, destinationPath string
	err := error.Error()
	for err != nil {
		fmt.Println("What is the path to the roster?")
	}
	fmt.Println("What is the path to the base report?")
	fmt.Scan(&templatePath)
	fmt.Println("What is the path of the output?")
	fmt.Scan(&destinationPath)

	
}
