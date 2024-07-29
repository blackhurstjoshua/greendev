package main

import (
	"archive/zip"
	"log"
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
func readXLSX(filepath string) *Worksheet {
	r, err := zip.OpenReader(filepath)
	if check(err) {
		panic(err.Error())
	}
	defer r.Close()

	// var sheet Worksheet

	// ? xlsx files are basically just a zip file LAME
	for _, f := range r.File {
		log.SetFlags(0)
		log.Println(f.Name)
	}

	return nil
}

func main() {
	filepath := "/Users/joshua/Downloads/3089.xlsx"
	s := readXLSX(filepath)

	log.Print(s)
}
