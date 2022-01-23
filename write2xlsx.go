package main

import (
	"fmt"
  "os"
	"github.com/xuri/excelize/v2"
)

func main() {

  f, err := excelize.OpenFile(os.Args[1])
    if err != nil {
        fmt.Println(err)
        return
    }
    defer func() {
        // Close the spreadsheet.
        if err := f.Close(); err != nil {
            fmt.Println(err)
        }
    }()
    // Get all the rows in the Sheet1.
    sheetName := f.GetSheetName(f.GetActiveSheetIndex())
    rows, err := f.GetRows(sheetName)
    if err != nil {
        fmt.Println(err)
        return
    }
    for x, row := range rows {
        for y, colCell := range row {
            sy,_:=excelize.CoordinatesToCellName(y+1, x+1)
            f.SetCellValue(sheetName, sy, colCell)
        }
    }
  f.SetCellValue(sheetName, os.Args[2], os.Args[3])
	f.Save()
  fmt.Println("OK")
}
