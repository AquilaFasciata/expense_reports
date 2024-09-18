package main

import "fmt"

func main() {
	var rosterPath, templatePath, destinationPath string
	fmt.Println("What is the path to the roster?")
	_, err := fmt.Scan(&rosterPath)
	for err != nil {
		fmt.Println("There was an error. You can try a different path, though!")
		_, err = fmt.Scan(&rosterPath)
	}
	_, err = fmt.Println("What is the path to the base report?")
	fmt.Scan(&templatePath)
	for err != nil {
		fmt.Println("There was an error. You can try a different path, though!")
		_, err = fmt.Scan(&templatePath)
	}
	fmt.Println("What is the path of the output?")
	_, err = fmt.Scan(&destinationPath)
	for err != nil {
		fmt.Println("There was an error. You can try a different path, though!")
		_, err = fmt.Scan(&rosterPath)
	}
}
