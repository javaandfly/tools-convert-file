package main

import (
	"bytes"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"
	"strings"

	"github.com/tealeg/xlsx"
	"github.com/urfave/cli"
)

const (
	fileFlag            = "file"
	fileShortFlag       = "f"
	fileIOFlag          = "output"
	fileIOShortFlag     = "o"
	fileSuffixShortFlag = "s"
	fileSuffixFlag      = "suffix"
)

func main() {
	cli.AppHelpTemplate = fmt.Sprintf(
		`%sExplanation: 
		This command line is used for file parsing and conversion
		`, cli.AppHelpTemplate)
	app := &cli.App{
		Flags: []cli.Flag{
			&cli.StringFlag{
				Name:     fileFlag + "," + fileShortFlag,
				Usage:    "file path",
				Required: true,
			},
			&cli.StringFlag{
				Name:  fileIOFlag + "," + fileIOShortFlag,
				Usage: "file io",
			},
			&cli.StringFlag{
				Name:  fileSuffixFlag + "," + fileSuffixShortFlag,
				Usage: "file suffix",
			},
		},
		Action: MainAction,
	}

	err := app.Run(os.Args)
	if err != nil {
		log.Panicln(err)
		return
	}
}

func MainAction(c *cli.Context) (err error) {
	filePath := c.String(fileFlag)
	fileIOPath := c.String(fileIOFlag)
	fileIOType := c.String(fileSuffixFlag)
	if fileIOPath == "" || filePath == "" || fileSuffixFlag == "" {
		err = fmt.Errorf("param error")
		return err
	}
	_, err = os.Stat(fileIOPath)
	if err != nil {
		err := os.Mkdir(fileIOPath, os.ModePerm)
		if err != nil {
			return err
		}
	}
	files := strings.Split(filePath, ",")
	for _, file := range files {

		fileType := path.Ext(file)

		switch {
		case fileType == ".xlsx", fileIOType == ".csv":
			strs := strings.Split(file, "/")
			fileName := strings.Split(strs[len(strs)-1], ".")
			path := fmt.Sprintf("%s/%s", fileIOPath, fileName[0])
			fmt.Println(fileIOType)
			err = convFileXlsxToCsv(file, path, fileIOType)

		default:
			err = fmt.Errorf("sorry not currently supported")
		}
		if err != nil {
			return err
		}

	}
	log.Println("convert success")
	return
}

// ConvFileExtension converts file to another file, different file extension
// is supported
func convFileXlsxToCsv(oldFileName, newFileName, fileType string) error {
	f, err := xlsx.OpenFile(oldFileName)
	if err != nil {
		err := fmt.Errorf("read file error %v", err)
		return err
	}
	_, err = os.Stat(newFileName)
	if err != nil {
		err := os.Mkdir(newFileName, os.ModePerm)
		if err != nil {
			return err
		}
	}

	for _, sheet := range f.Sheets {
		fileName := fmt.Sprintf("%s/%s%s", newFileName, sheet.Name, fileType)
		fmt.Println(newFileName, sheet.Name, fileType)
		err := parseSheetToCSV(sheet, fileName)
		if err != nil {
			err := fmt.Errorf("%s file is error is: %v", oldFileName, err)
			return err
		}
	}
	log.Printf("%s success", newFileName)
	return nil
}

func parseSheetToCSV(sheet *xlsx.Sheet, toFile string) (err error) {
	b := &bytes.Buffer{}
	rows := sheet.MaxRow
	for i := 0; i < rows; i++ {
		cols := sheet.MaxCol
		for j := 0; j < cols; j++ {
			cell := sheet.Cell(i, j)
			val := cell.Value
			// fmt.Println(val)
			b.WriteString(val)
			b.WriteString(",")
		}
		b.WriteString("\r\n")
	}
	//写入数据到文件
	err = ioutil.WriteFile(toFile, b.Bytes(), os.ModePerm)
	return
}
