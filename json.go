package main

import (
    "fmt"
    "io/ioutil"
    "os"
    "time"
    "encoding/json"
    "github.com/tealeg/xlsx"
)

var (
    targetExcel string = "./sitelist.xlsx"
    outputJSON  string = "./meta.json"
    sheetName string = "sitedata"
)


// Excelから取得するデータ
type pageData struct {
    Title       string `json:"title"`
    Url         string `json:"url"`
    Description string `json:"description"`
    Keywords    string `json:"keywords"`
    Ogtitle     string `json:"ogtitle"`
    Ogurl       string `json:"ogurl"`
    Ogimg       string `json:"ogimg"`
    Canonical   string `json:"canonical"`
}


func main() {

    // エクセルファイルをmapに変換する
    getExcelSheet(sheetName)
}

// エクセルファイルをmapに変換する
func getExcelSheet( aSheetName string ) {

    start := time.Now();

    // sliceの宣言
    sitedata := make([]pageData, 0)


    // エクセルファイルの取得
    excelFileName := targetExcel

    xlFile, err := xlsx.OpenFile(excelFileName)
    if err != nil {
        fmt.Println("エラー")
    }

    for _, sheet := range xlFile.Sheets {

        if aSheetName == sheet.Name {

            // TODO シートのデータ受け取ったら別関数に値ごと渡す

            for _, row := range sheet.Rows {

                data := &pageData{ }
                hasData := true

                for cellCnt, cell := range row.Cells {
                    // fmt.Printf("%s\n", cell.String())

                    if cell.String() == "" {
                        hasData = false
                        break
                    }

                    if cellCnt == 0 {
                        data.Title = cell.String()
                    } else if cellCnt == 1 {
                        data.Url = cell.String()
                    } else if cellCnt == 2 {
                        data.Description = cell.String()
                    } else if cellCnt == 3 {
                        data.Keywords = cell.String()
                    } else if cellCnt == 4 {
                        data.Ogtitle = cell.String()
                    } else if cellCnt == 5 {
                        data.Ogurl = cell.String()
                    } else if cellCnt == 6 {
                        data.Ogimg = cell.String()
                    } else if cellCnt == 7 {
                        data.Canonical = cell.String()
                    }

                }

                if data.Title == "" {
                    fmt.Println(data)
                    continue
                }

                if hasData {
                    sitedata = append(sitedata, *data)
                }

            }

            // TODO データの書き出しは別関数で実装
            output, _ := json.Marshal(sitedata)


            content := []byte(output)
            ioutil.WriteFile(outputJSON, content, os.ModePerm)


            break
        }

    }

    end := time.Now();
    fmt.Printf("%f秒\n",(end.Sub(start)).Seconds())

}