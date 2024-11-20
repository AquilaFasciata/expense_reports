package main

import (
	"errors"
	"fmt"
	"os"
	"slices"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"

	"github.com/xuri/excelize/v2"
)

type InputType int

const (
	FILE   InputType = iota
	FOLDER InputType = iota
)

func main() {
	prog := app.New()
	mainWindow := prog.NewWindow("Expense Report Builder")
	mainWindow.Resize(fyne.Size{Height: 412, Width: 720})

	templateRow, templatePathBox := inputRow(&mainWindow, "Template: ", FILE)
	rosterRow, rosterPathBox := inputRow(&mainWindow, "Roster: ", FILE)
	outputRow, outputPathBox := inputRow(&mainWindow, "Output: ", FOLDER)
	runButton := widget.NewButton("Create Expense Reports", func() {
		err := business(rosterPathBox.Text, templatePathBox.Text, outputPathBox.Text)
		if err != nil {
			dialog.ShowError(err, mainWindow)
		}
	})

	mainWindow.SetContent(container.NewStack(
		container.NewVBox(
			templateRow, rosterRow, outputRow, runButton,
		),
	))

	mainWindow.ShowAndRun()
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

func business(rosterPath, templatePath, destinationPath string) error {
	// scanner := bufio.NewReader(os.Stdin)
	// fmt.Println("What is the path to the roster?")
	// rosterPath, err := scanner.ReadString('\n')
	// rosterPath = strings.TrimSpace(rosterPath)
	// for err != nil {
	// 	fmt.Println(err)
	// 	return err
	// }

	roster, err := excelize.OpenFile(strings.TrimSpace(rosterPath))
	if err != nil {
		fmt.Println("Error reading roster: ", err)
		return err
	}

	// fmt.Println("What is the path to the base report?")
	// templatePath, err = scanner.ReadString('\n')
	// templatePath = strings.TrimSpace(templatePath)
	// templatePath = strings.Trim(templatePath, "\"")
	// for err != nil {
	// 	fmt.Println(err)
	// 	return err
	// }

	template, err := excelize.OpenFile(strings.TrimSpace(templatePath))
	if err != nil {
		fmt.Println("Error reading template: ", err)
		return err
	}

	_, err = os.ReadDir(destinationPath)
	if err != nil {
		fmt.Println("Error opening directory: ", err, "\n\n")
		return err
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
	return nil
}

func inputRow(parentWindow *fyne.Window, label string, openType InputType) (*fyne.Container, *widget.Entry) {
	input := widget.NewEntry()
	if openType == FILE {
		input.SetPlaceHolder("C:\\Users\\me\\file.xlsx")
	} else {
		input.SetPlaceHolder("C:\\Users\\me\\folder")
	}
	input.Resize(fyne.NewSize(400, input.Size().Height))

	inputLabel := widget.NewLabel(label)
	var containy *fyne.Container
	if openType == FILE {
		containy = container.NewGridWithColumns(3, inputLabel, input, widget.NewButton("Browse", func() { fileDialog(*parentWindow, input) }))
	} else {
		containy = container.NewGridWithColumns(3, inputLabel, input, widget.NewButton("Browse", func() { folderDialog(*parentWindow, input) }))
	}
	return containy, input
}

func fileDialog(parentWindow fyne.Window, box *widget.Entry) error {
	xlExtensions := []string{".xlsx", ".xls", ".xlsm"}
	var anyError error = nil
	dialog.ShowFileOpen(func(reader fyne.URIReadCloser, err error) {
		if err != nil {
			anyError = err
			fmt.Println(err)
			dialog.ShowError(err, parentWindow)
			return
		}
		if !slices.Contains(xlExtensions, reader.URI().Extension()) {
			anyError = errors.New("Selected file is not a valid Excel file!")
			dialog.ShowError(anyError, parentWindow)
			return
		}
		fmt.Println(reader.URI().Extension())
		box.Text = ""
		box.Text = reader.URI().Path()
		box.Refresh()
	}, parentWindow)

	return anyError
}

func folderDialog(parentWindow fyne.Window, box *widget.Entry) error {
	var anyError error = nil
	dialog.ShowFolderOpen(func(reader fyne.ListableURI, err error) {
		if err != nil {
			anyError = err
			fmt.Println(err)
			dialog.ShowError(err, parentWindow)
			return
		}
		box.Text = ""
		box.Text = reader.Path()
		box.Refresh()
	}, parentWindow)

	return anyError
}
