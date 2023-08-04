package main

import (
	"bytes"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"

	"github.com/tealeg/xlsx"
	"github.com/urfave/cli"
)

const (
	fileFlag        = "file"
	fileShortFlag   = "f"
	fileIOFlag      = "output"
	fileIOShortFlag = "o"
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
	if fileIOPath == "" || filePath == "" {
		err = fmt.Errorf("param error")
		return err
	}
	fileType := path.Ext(filePath)
	fileIOType := path.Ext(fileIOPath)
	switch {
	case fileType == ".xlsx", fileIOType == ".csv":
		err = convFileXlsxToCsv(fileIOPath, filePath)

	default:
		err = fmt.Errorf("sorry not currently supported")
	}
	if err != nil {
		return err
	}
	log.Println("convert success")
	return
}

// ConvFileExtension converts file to another file, different file extension
// is supported
func convFileXlsxToCsv(newFileName string, oldFileName string) error {
	f, err := xlsx.OpenFile(oldFileName)
	if err != nil {
		err := fmt.Errorf("read file error %v", err)
		return err
	}
	return parseSheetToCSV(f.Sheets[0], newFileName)
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
			b.WriteString("\t")
		}
		b.WriteString("\r\n")
	}
	//写入数据到文件
	err = ioutil.WriteFile(toFile, b.Bytes(), os.ModePerm)
	return
}
