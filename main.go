/**
 * Seiichi Ariga <seiichi.ariga@gmail.com> (c) 2022
 * クラブ対抗戦の領収書を作成する
 * Excelテンプレートに宛先、金額、適用を流し込みます
 */

package main

import (
	"fmt"
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

	for _, receipt := range receipts {
		fmt.Println(receipt)
	}

	// 出力ファイルを書き込む
	err = receipt.WriteReceipts(receipts)
	if err != nil {
		log.Panic(err)
	}
}
