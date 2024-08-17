package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"strconv"
)

// read the excel file without using external package
type SharedStringItem struct {
	T string `xml:"t"`
}

type SharedStrings struct {
	SI []SharedStringItem `xml:"si"`
}

type Cell struct {
	Type  string `xml:"t,attr"`
	Value string `xml:"v"`
}

type Row struct {
	Cells []Cell `xml:"c"`
}

type Sheet struct {
	Rows []Row `xml:"row"`
}

type Worksheet struct {
	Sheet Sheet `xml:"sheetData"`
	SS    *SharedStrings
}

func check(e error) bool {
	return e != nil
}
func readXLSX(filepath string, sheet string) *Worksheet {
	r, err := zip.OpenReader(filepath)
	if check(err) {
		panic(err.Error())
	}
	defer r.Close()

	var ws Worksheet

	// ? xlsx files are basically just a zip file - LAME
	for _, f := range r.File {
		fmt.Println(f.Name)
		if f.Name == "xl/worksheets/sheet1.xml" {
			rc, err := f.Open()
			if check(err) {
				panic("can't open sheet: " + err.Error())
			}
			defer rc.Close()

			decoder := xml.NewDecoder(rc)
			err = decoder.Decode(&ws)
			if check(err) {
				panic("could not decode sheet: " + err.Error())
			}
		} else if f.Name == "xl/sharedStrings.xml" {
			rc, err := f.Open()
			if check(err) {
				panic("can't open sharedStrings sheet: " + err.Error())
			}
			defer rc.Close()

			decoder := xml.NewDecoder(rc)
			err = decoder.Decode(&ws.SS)
			if check(err) {
				panic("could not decode sharedStrings sheet: " + err.Error())
			}
		}
	}

	return &ws
}

func (ws *Worksheet) headerRowIndex() (int, error) {
	data := ws.Sheet.Rows
	for i := 0; i < len(data); i++ {
		if len(data[i].Cells) > 0 {
			// potential header found
			return i, nil
		}
	}
	return 0, fmt.Errorf("unable to locate potential header row")
}

func (ws *Worksheet) suggestHeader(idx int) ([]string, error) {
	header := ws.Sheet.Rows[idx]
	headerStr := []string{}

	for _, cell := range header.Cells {
		if cell.Type == "s" {
			idx, _ := strconv.Atoi(cell.Value)
			headerStr = append(headerStr, ws.SS.SI[idx].T)
		} else {
			return []string{}, fmt.Errorf("all header names must be strings")
		}
	}

	return headerStr, nil
}

func main() {
	filepath := "/Users/joshua/Downloads/Copy of 3089.xlsx"
	ws := readXLSX(filepath, "Sheet1")

	headerIdx, err := ws.headerRowIndex()
	if check(err) {
		panic(err)
	}
	fmt.Println(headerIdx)

	fmt.Println(ws.suggestHeader(0))
	// fmt.Printf("Row length: %v\n", len(ws.Sheet.Rows[1].Cells))
	// for _, row := range ws.Sheet.Rows {
	// 	for _, cell := range row.Cells {
	// 		var value string
	// 		if cell.Type == "s" {
	// 			index, _ := strconv.Atoi(cell.Value)
	// 			value = ss.SI[index].T
	// 		} else {
	// 			value = cell.Value
	// 		}
	// 		fmt.Print(value, "\t")
	// 	}
	// }
}
