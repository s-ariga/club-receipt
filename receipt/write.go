/**
 * Seiichi Ariga <seiichi.ariga@gmail.com>
 */

package receipt

import (
	"errors"
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

const (
	TEMPLATE   = "./template/receipt-template.xlsx"
	OUTPUT_DIR = "./output/"
	FILE_EXT   = "様領収書.xlsx"
	SHEET_NAME = "領収書"
)

// 同じファイル名が存在するか調べる
func checkFileExists(filePath string) bool {
	_, err := os.Stat(filePath)
	return !errors.Is(err, os.ErrNotExist)
}

// OUTPUT_DIR内のファイルを「全部」消す
func cleanOutputDir() error {
	// TODO: たぶん、下の行が問題
	//err := os.RemoveAll(OUTPUT_DIR)
	//if err != nil {
	//	return err
	//}

	_, err = os.Stat(OUTPUT_DIR)
	// 出力ディレクトリが無かったら、作る
	if os.IsNotExist(err) {
		err = os.Mkdir(OUTPUT_DIR, 0755)
		if err != nil {
			log.Fatal(err)
		}
	}

	return nil
}

// 同じファイル名をつけないので、実行前にファイル全部消して作り直す
func WriteReceipts(receipts []Receipt) error {

	// TODO: ディレクトリが見つからずエラーになるときがある。
	// Google Driveがネットワークアクセスできない場合？

	err := cleanOutputDir()
	if err != nil {
		log.Panic(err)
	}
	// receiptsはレシートのリストなので、1つずつファイルに書き込んでいく
	for _, receipt := range receipts {
		fileName := OUTPUT_DIR + receipt.Name + FILE_EXT
		f, err := excelize.OpenFile(TEMPLATE)
		if err != nil {
			log.Panic(err)
		}
		defer f.Close()

		f.SetCellValue(SHEET_NAME, "A7", receipt.Name)     // 宛名
		f.SetCellValue(SHEET_NAME, "A21", receipt.Summary) // 摘要
		f.SetCellValue(SHEET_NAME, "E21", receipt.Price)   // 単価
		f.SetCellValue(SHEET_NAME, "E30", receipt.Price)   // 合計金額
		f.SetCellValue(SHEET_NAME, "B13", receipt.Price)   // 領収金額

		// 同じファイル名が存在したら、"-1", "-2" ... をつけていく
		i := 1
		for checkFileExists(fileName) {
			num := fmt.Sprint(i)
			fileName = OUTPUT_DIR + receipt.Name + "-" + num + FILE_EXT
			i += 1
		}
		err = f.SaveAs(fileName)
		if err != nil {
			log.Panic(err)
		}
		f.Close()
	}
	return nil
}
