/**
 * Seiichi Ariga <seiichi.ariga@gmail.com> (c) 2022
 */

package main

import (
	"log"

	"github.com/s-ariga/club-receipt/receipt"
)

// 入力ファイル名
const (
	INPUT = "./input/receipt-list.xlsx"
)

func main() {
	// 入力ファイルを読み込む
	receipts, err := receipt.ReadReceipts(INPUT)
	if err != nil {
		log.Panic(err)
	}

	// 出力ファイルを書き込む
	err = receipt.WriteReceipts(receipts)
	if err != nil {
		log.Panic(err)
	}
}
