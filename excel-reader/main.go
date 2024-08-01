package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"strings"
)

// read the excel file without using external package
type Cell struct {
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
		if strings.Contains(f.Name, "xl/worksheets/"+sheet) {
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
		}
	}

	return &ws
}

func main() {
	filepath := "/Users/joshua/Downloads/Copy of 3089.xlsx"
	s := readXLSX(filepath, "sheet1")

	fmt.Printf("Row length: %v\n", len(s.Sheet.Rows[1].Cells))
	for _, r := range s.Sheet.Rows {
		fmt.Printf("%v\n", string(r.Cells[0].Value))
	}
}
