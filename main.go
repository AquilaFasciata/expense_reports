package main

import (
	"bufio"
	"fmt"
	"os"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"

	"github.com/xuri/excelize/v2"
)

func main() {
	var rosterPath, templatePath, destinationPath string

	prog := app.New()
	mainWindow := prog.NewWindow("Expense Report Builder")

	mainWindow.SetContent(container.NewVBox())

	business(rosterPath, templatePath, destinationPath)
	// TODO Build ui
	// TODO Verify additions from Github
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

func business(rosterPath, templatePath, destinationPath string) {
	scanner := bufio.NewReader(os.Stdin)
	fmt.Println("What is the path to the roster?")
	rosterPath, err := scanner.ReadString('\n')
	rosterPath = strings.TrimSpace(rosterPath)
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
	templatePath, err = scanner.ReadString('\n')
	templatePath = strings.TrimSpace(templatePath)
	templatePath = strings.Trim(templatePath, "\"")
	for err != nil {
		fmt.Println(err)
		return
	}

	template, err := excelize.OpenFile(strings.TrimSpace(templatePath))
	if err != nil {
		fmt.Println("Error reading template: ", err)
		return
	}

	for {
		fmt.Println("What is the path of the output?")
		destinationPath, err = scanner.ReadString('\n')
		destinationPath = strings.TrimSpace(destinationPath)
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

	roster_cols, err := roster.GetCols("Sheet1")
	err_check(err)

	today := time.Now()
	for i, name := range roster_cols[1] {
		if name == "" {
			continue
		}
		template.SetCellStr("Expense Report Template", "D7", name)
		template.SetCellStr("Expense Report Template", "D6", roster_cols[0][i])
		template.SetCellStr("Expense Report Template", "C13", "Welcome to Mike's")
		template.SetCellStr("Expense Report Template", "E13", "Yes")
		template.SetCellStr("Expense Report Template", "B13", today.Format("01/02/2006"))

		for j := 1; j < len(locations); j++ {
			location := get_loc_num(locations[j][0])
			roster_loc := roster_cols[2][i]
			if strings.Compare(location, roster_loc) == 0 {
				template.SetCellStr("Expense Report Template", "D8", locations[j][0])
				template.SetCellStr("Expense Report Template", "F13", locations[j][0])

				break
			}
		}

		template.UpdateLinkedValue()
		template.SaveAs(destinationPath + "/" + name + ".xlsm")
		template, _ = excelize.OpenFile(templatePath)
	}
}

func fileOpenRow(parentWindow fyne.Window, label string, pathBuff *string) (*fyne.Container, *widget.Entry) {
	input := widget.NewEntry()
	input.SetPlaceHolder("C:\\Users\\me\\file.xlsx")
	containy := container.NewHBox(widget.NewLabel(label), input, widget.NewButton("Browse", func() { fileDialog(parentWindow, pathBuff) }))
	return containy, input
}

func fileDialog(parentWindow fyne.Window, pathBuff *string) {
	dialog.ShowFileOpen(func(reader fyne.URIReadCloser, err error) {
		if err != nil {
			fmt.Println(err)
			os.Exit(1)
		}
		*pathBuff = reader.URI().Path()
	}, parentWindow)
}
