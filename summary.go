// must set environment variable GO_EXTLINK_ENABLED=0
package main

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/andlabs/ui"
	_ "github.com/andlabs/ui/winmanifest"

	"github.com/360EntSecGroup-Skylar/excelize"

	"os"
)

var (
	label    *ui.Label
	okButton *ui.Button
	result   string
	err      error
)

func main() {
	summaryFile := "Team Results Summary.xlsx"
	if len(os.Args) == 2 {
		summaryFile = os.Args[1]
	}
	_, err := os.Stat(summaryFile)
	if err != nil {
		fmt.Println("Cannot find summary file: " + summaryFile)
		result = "Cannot find summary file: " + summaryFile
		ui.Main(setupUI)
		os.Exit(0)
	}

	err = summary(summaryFile)
	if err != nil {
		result = err.Error()
	}
	ui.Main(setupUI)

}

func summary(summaryFile string) error {
	summarySheet := "Sheet1"

	var xlsx *excelize.File

	xlsx, err = excelize.OpenFile(summaryFile)
	if err != nil {

		return err

	}

	//		xlsx.SetCellValue("Sheet1", "A"+strconv.Itoa(i), pkt.String)
	sheetNames := strings.Split(xlsx.GetCellValue("configuration", "B1"), ",")
	for i, v := range sheetNames {
		sheetNames[i] = strings.TrimSpace(v)
	}

	columns := strings.Split(xlsx.GetCellValue("configuration", "B2"), ",")
	for i, v := range columns {
		columns[i] = strings.TrimSpace(v)
	}

	refCell := strings.TrimSpace(xlsx.GetCellValue("configuration", "B3"))

	rows := xlsx.GetRows(summarySheet)

	type nameRow struct {
		name string
		row  string
	}

	var nameRows []nameRow

	for i, v := range rows {
		if i == 0 {
			continue
		}
		if strings.TrimSpace(v[0]) != "" {
			nameRows = append(nameRows, nameRow{name: strings.TrimSpace(v[0]), row: strconv.Itoa(i + 1)})

		}

	}

	formsFound := 0
	for _, v := range nameRows {
		xlsx2, err2 := excelize.OpenFile(v.name + ".xlsx")
		if err2 != nil {
			for _, vv := range columns {

				xlsx.SetCellValue(summarySheet, vv+v.row, "")
			}
			continue
		}
		formsFound++
		for j, vv := range columns {

			f, err := strconv.ParseFloat(xlsx2.GetCellValue(sheetNames[j], refCell), 64)
			if err != nil {
				xlsx.SetCellValue(summarySheet, vv+v.row, "")
			} else {
				xlsx.SetCellValue(summarySheet, vv+v.row, f)
			}
		}

	}

	result = "Forms found/expected: " + strconv.Itoa(formsFound) + "/" + strconv.Itoa(len(nameRows))
	xlsx.Save()
	return nil
}

func setupUI() {

	mainwin := ui.NewWindow("Summary Tool", 150, 50, false)

	mainwin.SetMargined(true)
	mainwin.OnClosing(func(*ui.Window) bool {
		mainwin.Destroy()
		ui.Quit()
		return false
	})
	ui.OnShouldQuit(func() bool {
		mainwin.Destroy()
		return true
	})

	hbox := ui.NewHorizontalBox()
	hbox.SetPadded(true)

	mainwin.SetChild(hbox)

	vbox := ui.NewVerticalBox()
	vbox.SetPadded(true)

	hbox.Append(vbox, false)

	label = ui.NewLabel(result)

	vbox.Append(label, false)
	okButton = ui.NewButton("OK")
	okButton.OnClicked(func(*ui.Button) {
		mainwin.Destroy()
		ui.Quit()

	})
	vbox.Append(okButton, false)

	mainwin.Show()
}
