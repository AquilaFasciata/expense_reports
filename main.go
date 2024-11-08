package main

import (
	"bufio"
	"fmt"
	"image/color"
	"io"
	"log"
	"os"
	"strings"
	"time"

	"gioui.org/app"
	"gioui.org/layout"
	"gioui.org/op"
	"gioui.org/text"
	"gioui.org/unit"
	"gioui.org/widget"
	"gioui.org/widget/material"

	// "gioui.org/x/explorer"
	"github.com/xuri/excelize/v2"
)

func main() {
	var rosterPath, templatePath, destinationPath string

	go func() {
		window := new(app.Window)
		err := run(window)
		if err != nil {
			log.Fatal(err)
		}
		os.Exit(1)
	}()

	business()

	// TODO Build ui
	// TODO Verify additions from Github
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

func run(window *app.Window) error {
	theme := material.NewTheme()

	button := widget.Clickable{}
	var ops op.Ops
	for {
		switch evnt := window.Event().(type) {
		case app.DestroyEvent:
			return evnt.Err
		case app.FrameEvent:
			gtx := app.NewContext(&ops, evnt)

			title := material.H1(theme, "Hello, GUI!")
			button := material.Button(theme, &button, "Test")

			maroon := color.NRGBA{R: 127, G: 0, B: 0, A: 255}
			title.Color = maroon

			title.Alignment = text.Middle

			layout.Flex{
				Axis:    layout.Vertical,
				Spacing: layout.SpaceStart,
			}.Layout(gtx,
				layout.Rigid(func(gtx layout.Context) layout.Dimensions {
					return title.Layout(gtx)
				}),
				layout.Rigid(func(gtx layout.Context) layout.Dimensions {
					return layout.Spacer{Height: unit.Dp(75)}.Layout(gtx)
				}),
				layout.Rigid(func(gtx layout.Context) layout.Dimensions {
					return button.Layout(gtx)
				}))

			evnt.Frame(gtx.Ops)
		}
	}
}

func file_input(gtx layout.Context, label string, theme material.Theme) (layout.FlexChild, widget.Editor) {
	input_box := widget.Editor
	returned_layout := layout.Flex{
		Axis: layout.Horizontal,
		Spacing: layout.SpaceStart,
	}.Layout(gtx, 
		layout.Rigid(func(gtx layout.Context) layout.Dimensions {
			return material.Label(&theme, theme.TextSize, label + ":")
		}),
		layout.Rigid(func(gtx layout.Context) layout.Dimensions {
			return material.Editor(&theme, )
		})
	)

	return (returned_layout, input_box)
}
