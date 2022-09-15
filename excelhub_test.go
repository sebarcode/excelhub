package excelhub

import (
	"os"
	"path/filepath"
	"testing"
	"time"

	"github.com/sebarcode/codekit"
	"github.com/smartystreets/goconvey/convey"
)

func TestIncreaseRow(t *testing.T) {
	TestData := []struct {
		Name   string
		Row    string
		Result string
	}{
		{"Simple Next 1", "A", "B"},
		{"Simple Next 2", "D", "E"},
		{"1 Char to 2 Char", "Z", "AA"},
		{"2 Char - same level", "AD", "AE"},
		{"2 Char - different level 1", "AZ", "BA"},
		{"2 Char - different level 2", "CZ", "DA"},
		{"2 Char to 3 Char", "ZZ", "AAA"},
		{"3 Char - case 1", "AAC", "AAD"},
		{"3 Char - case 2", "ABC", "ABD"},
		{"3 Char - level 1", "ABZ", "ACA"},
		{"3 Char - level 2", "CZZ", "DAA"},
	}

	convey.Convey("increase row", t, func() {
		for _, td := range TestData {
			convey.Convey(td.Name, func() {
				res, _ := increaseRow(td.Row)
				convey.So(res, convey.ShouldEqual, td.Result)
			})
		}
	})
}

func TestNextCell(t *testing.T) {
	TestData := []struct {
		Name      string
		Direction string
		Cell      string
		Result    string
	}{
		{"H Simple", "H", "A1", "B1"},
		{"H Naik Level", "H", "Z1", "AA1"},
		{"H Level 2a", "H", "BA1", "BB1"},
		{"H Level 2b", "H", "BZ1", "CA1"},
		{"H Naik level 3", "H", "ZZ1", "AAA1"},
		{"V - 1", "V", "AA1", "AA2"},
		{"V - 2", "V", "DE1", "DE2"},
	}

	convey.Convey("next cell", t, func() {
		for _, td := range TestData {
			convey.Convey(td.Name, func() {
				res := nextCell(td.Cell, td.Direction)
				convey.So(res, convey.ShouldEqual, td.Result)
			})
		}
	})
}

func TestExportToExcel(t *testing.T) {
	wd, _ := os.Getwd()
	outputFile := "output.xlsx"
	exampleFolder := filepath.Join(wd, "example")
	//destFile := filepath.Join(exampleFolder, outputFile)
	defer func() {
		//os.Remove(destFile)
	}()

	opts := &Options{
		TemplateFile: filepath.Join(exampleFolder, "template1.xlsx"),
		SheetName:    "Sheet1",
		OutputFolder: exampleFolder,
		StartCell:    "A6",
		EndCell:      "H6",
		CellMaps: []CellMap{
			{Cell: "B2", MapTo: "Rig"},
			{Cell: "B3", MapTo: "Well"},
			{Cell: "E2", MapTo: "Field"},
			{Cell: "E3", MapTo: "Location"},
			{Cell: "H2", MapTo: "Start"},
			{Cell: "H3", MapTo: "End"},
		},
	}

	convey.Convey("export to excel", t, func() {
		headerData := codekit.M{
			"Rig": "AL 101HSA", "Well": "AL 101", "Location": "Tx",
			"Start": time.Date(2022, 9, 20, 0, 0, 0, 0, time.Now().Location()),
			"End":   time.Date(2022, 9, 20, 2, 0, 0, 0, time.Now().Location()),
		}

		startDate := time.Date(2022, 9, 20, 0, 0, 0, 0, time.Now().Location())
		rowData := []codekit.M{
			codekit.M{}.Set("Time", startDate).Set("Op", "Drilling").Set("Lith", "Rock").Set("Depth", 201.80).Set("Pressure", 187.60),
			codekit.M{}.Set("Time", startDate.Add(1*time.Second)).
				Set("Op", "Drilling").Set("Lith", "Water").
				Set("Depth", 201.80).Set("Pressure", 187.60),
			codekit.M{}.Set("Time", startDate.Add(2*time.Second)).
				Set("Op", "Drilling").Set("Lith", "Water").Set("Temp", 87.52).
				Set("Depth", 202.75).Set("Pressure", 151.34),
		}
		e := ExportToExcel(outputFile, opts, headerData, rowData)
		convey.So(e, convey.ShouldBeNil)
	})
}
