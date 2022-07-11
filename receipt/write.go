/**
 * Seiichi Ariga <seiichi.ariga@gmail.com>
 */

package receipt

import (
	"github.com/xuri/excelize/v2"
)

const (
	TEMPLATE   = "./template/receipt-template.xlsx"
	OUTPUT_DIR = "./output/"
	FILE_EXT   = "様領収書.xlsx"
	SHEET_NAME = "領収書"
)

func WriteReceipts(receipts []Receipt) error {
	// receiptsはレシートのリストなので、1つずつファイルに書き込んでいく

	for _, receipt := range receipts {
		fileName := OUTPUT_DIR + receipt.Name + FILE_EXT
		f, err := excelize.OpenFile(TEMPLATE)
		if err != nil {
			return err
		}
		defer f.Close()

		f.SetCellValue(SHEET_NAME, "A7", receipt.Name)     // 宛名
		f.SetCellValue(SHEET_NAME, "A21", receipt.Summary) // 摘要
		f.SetCellValue(SHEET_NAME, "E21", receipt.Price)   // 単価
		f.SetCellValue(SHEET_NAME, "E30", receipt.Price)   // 合計金額
		f.SetCellValue(SHEET_NAME, "B13", receipt.Price)   // 領収金額

		err = f.SaveAs(fileName)
		if err != nil {
			return err
		}
		f.Close()
	}
	return nil
}
