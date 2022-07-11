/**
 * Seiichi Ariga <seiichi.ariga@gmail.com>
 */

package receipt

import (
	"strconv"

	"github.com/xuri/excelize/v2"
)

type Receipt struct {
	Name    string
	Price   int
	Summary string
}

const (
	SHEET = "フォームの回答 1"
)

func ReadReceipts(fileName string) ([]Receipt, error) {

	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	rows, err := f.GetRows(SHEET)
	if err != nil {
		return nil, err
	}

	receipts := make([]Receipt, 0)

	for i, row := range rows {
		if i == 0 {
			continue
		}
		price, err := strconv.Atoi(row[3])
		if err != nil {
			return nil, err
		}
		receipt := Receipt{row[2], price, row[4]}
		receipts = append(receipts, receipt)
	}
	return receipts, nil
}
