/*
 * @file
 * @brief           whatis
 * @note            how it work
 * @author          zhangjiayi
 * @date            2021-02-21 22:24:22
 * @version         v0.1
 * @copyright       Copyright (c) 2020-2050  zhangjiayi
 * @par             LastEdit
 * @LastEditTime    2021-02-22 00:21:27
 * @LastEditors     zhangjiayi
 * @FilePath        /excel/excel.go
 */

package main

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	readExcel, err := excelize.OpenFile("./read.xlsx")
	if err != nil {
		fmt.Println(err, "line 26")
		return
	}
	writeExcel, err := excelize.OpenFile("./write.xlsx")
	if err != nil {
		fmt.Println(err, "line 31")
		return
	}

	// Get all the rows in the Sheet1.
	readRows := readExcel.GetRows("Sheet1")
	writeRows := writeExcel.GetRows("古东")
	countW := 1
	for _, row := range writeRows {
		countR := 1
		for _, rowR := range readRows {
			if rowR[6] == row[4] && rowR[6] != "" {
				fmt.Print(countW, "\t", countR, "\t", "AH"+strconv.Itoa(countW), "\t", row[4], "\t")
				fmt.Println(rowR[59])
				writeExcel.SetCellStr("古东", "AH"+strconv.Itoa(countW), rowR[59])
			}
			countR++
		}
		countW++
	}
	err = writeExcel.Save()
	if err != nil {
		fmt.Println(err)
	}
}
