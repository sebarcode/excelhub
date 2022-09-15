package excelhub

import (
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/sebarcode/codekit"
	"github.com/xuri/excelize/v2"
)

type CellMap struct {
	Cell  string
	MapTo string
}

type Options struct {
	TemplateFile string
	SheetName    string
	OutputFolder string
	CellMaps     []CellMap
	StartCell    string
	EndCell      string
}

func ExportToExcel(outputName string, opts *Options, headerData codekit.M, data []codekit.M) error {
	if opts == nil {
		opts = new(Options)
	}

	//-- copy template to new file
	destFilePath := filepath.Join(opts.OutputFolder, outputName)
	if e := func() error {
		if s, err := os.Stat(opts.OutputFolder); err != nil {
			return fmt.Errorf("error reading output folder %s. %s", opts.OutputFolder, err.Error())
		} else if !s.IsDir() {
			return fmt.Errorf("error reading output folder %s. is not a directory", opts.OutputFolder)
		}

		source, err := os.Open(opts.TemplateFile)
		if err != nil {
			return fmt.Errorf("error reading file %s. %s", opts.TemplateFile, err.Error())
		}
		defer source.Close()

		dest, err := os.Create(destFilePath)
		if err != nil {
			return fmt.Errorf("fail to create the output file. %s", err)
		}
		defer dest.Close()

		_, err = io.Copy(dest, source)
		if err != nil {
			return fmt.Errorf("fail to copy template file to destination file. %s", err.Error())
		}

		return nil
	}(); e != nil {
		return e
	}

	//-- open excelize session
	f, err := excelize.OpenFile(destFilePath)
	if err != nil {
		return fmt.Errorf("fail to create excel session. %s", err.Error())
	}
	defer f.Close()
	if opts.SheetName == "" {
		opts.SheetName = "Sheet1"
	}

	//-- processing header
	for _, hdr := range opts.CellMaps {
		if err := f.SetCellValue(opts.SheetName, hdr.Cell, headerData[hdr.MapTo]); err != nil {
			return fmt.Errorf("fail to write cell %s. %s", hdr.Cell, err.Error())
		}
	}

	//-- get direction, direction is the direction data will be inserted
	direction := ""
	r1, _, _ := RowCol(opts.StartCell)
	r2, _, _ := RowCol(opts.EndCell)
	if r1 == r2 {
		direction = "H"
	} else {
		direction = "V"
	}

	//-- get data field to be inserted
	attrs := ReadCells(f, opts.SheetName, opts.StartCell, opts.EndCell, swapDirection(direction))
	if len(attrs) == 0 {
		return fmt.Errorf("invalid attribute setting")
	}
	attrNames := []string{}
	for _, attr := range attrs {
		attrNames = append(attrNames, attr.(string))
	}

	//-- processing data
	writtenCell := opts.StartCell
	for index, record := range data {
		if e := WriteData(f, opts.SheetName, attrNames, writtenCell, swapDirection(direction), record); e != nil {
			return fmt.Errorf("fail to write data %d. %s", index, e.Error())
		}
		writtenCell = nextCell(writtenCell, direction)
	}

	if e := f.Save(); e != nil {
		return fmt.Errorf("fail to save changes. %s", e.Error())
	}

	return nil
}

func WriteData(f *excelize.File, sheetName string, attrs []string, cellStart string, direction string, data codekit.M) error {
	currentCell := cellStart
	for _, attr := range attrs {
		f.SetCellValue(sheetName, currentCell, data[attr])
		currentCell = nextCell(currentCell, direction)
	}
	return nil
}

func swapDirection(dir string) string {
	if dir == "H" {
		return "V"
	}
	return "H"
}

func WriteCell(f *excelize.File, sheet, cell string, v interface{}) error {
	return f.SetCellValue(sheet, cell, v)
}

func nextCell(cell, direction string) string {
	r, c, _ := RowCol(cell)
	if direction == "H" {
		r, _ = increaseRow(r)
	} else {
		c++
	}
	newCell := fmt.Sprintf("%s%d", r, c)
	return newCell
}

func increaseRow(row string) (string, bool) {
	rowLen := len(row)

	newRow := ""
	increase := true
	for i := 0; i < rowLen; i++ {
		charIndex := rowLen - i - 1
		c := rune(row[charIndex])
		if increase {
			//cStr := string(c)
			//cStr, n := increaseRow(cStr)
			cStr := ""
			if c == 'Z' {
				cStr = "A"
				increase = true
			} else {
				c++
				cStr = string(c)
				increase = false
			}
			newRow = cStr + newRow
		} else {
			newRow = string(c) + newRow
		}
	}
	if increase {
		newRow = "A" + newRow
	}

	return newRow, false
}

func RowCol(cell string) (string, int, error) {
	row := ""
	col := 0

	for _, c := range cell {
		if c >= 'A' && c <= 'Z' {
			row += string(rune(c))
		} else {
			break
		}
	}

	colStr := strings.Replace(cell, row, "", -1)
	col, err := strconv.Atoi(colStr)

	return row, col, err
}

func ReadCells(f *excelize.File, shetName, startCell, endCell, direction string) []interface{} {
	res := []interface{}{}

	currentCell := startCell
	for {
		if currentCell == endCell {
			break
		}

		str, err := f.GetCellValue(shetName, currentCell)
		if err != nil {
			res = append(res, "")
		} else {
			tp, _ := f.GetCellType(shetName, currentCell)
			switch tp {
			case excelize.CellTypeBool:
				res = append(res, str == "true")

			case excelize.CellTypeNumber:
				num, _ := codekit.StringToFloat(str)
				res = append(res, num)

			case excelize.CellTypeDate:
				res = append(res, str)

			default:
				res = append(res, str)
			}
		}

		currentCell = nextCell(currentCell, direction)
	}
	return res
}
