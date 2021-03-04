/*
 * @file
 * @brief           whatis
 * @note            how it work
 * @author          zhangjiayi
 * @date            2021-02-21 22:24:22
 * @version         v0.1
 * @copyright       Copyright (c) 2020-2050  zhangjiayi
 * @par             LastEdit
 * @LastEditTime    2021-03-05 07:46:56
 * @LastEditors     zhangjiayi
 * @FilePath        /excelmatch/excelMatch.go
 */

package main

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	//源文件名输入
	var inputFile string = ""
	for {
		fmt.Print("请输入要匹配的源文件:")
		n, err := fmt.Scanln(&inputFile)
		if n != 1 {
			fmt.Println("输入参数数量不对")
			continue
		}
		if err != nil {
			fmt.Println(err)
			continue
		}
		break
	}
	readExcel, err := excelize.OpenFile(inputFile)
	if err != nil {
		fmt.Println(err)
		return
	}

	//源文件工作表名输入
	inSheet := ""
	for {
		fmt.Print("请输入要匹配的源文件工作表名:")
		n, err := fmt.Scanln(&inSheet)
		if n != 1 {
			fmt.Println("输入参数数量不对")
			continue
		}
		if err != nil {
			fmt.Println(err)
			continue
		}
		break
	}
	readRows := readExcel.GetRows(inSheet)

	//源文件要匹配的列输入
	inMatchCol := ""
	for {
		fmt.Print("请输入要匹配的源文件的列名:")
		n, err := fmt.Scanln(&inMatchCol)
		if n != 1 {
			fmt.Println("输入参数数量不对")
			continue
		}
		if err != nil {
			fmt.Println(err)
			continue
		}
		break
	}

	//源文件要复制的列输入
	inputCol := ""
	for {
		fmt.Print("请输入源文件要复制的列名:")
		n, err := fmt.Scanln(&inputCol)
		if n != 1 {
			fmt.Println("输入参数数量不对")
			continue
		}
		if err != nil {
			fmt.Println(err)
			continue
		}
		break
	}
	inColNum := excelize.TitleToNumber(inputCol)

	//目标文件名输入
	var outputFile string = ""
	for {
		fmt.Print("请输入要匹配的目标文件:")
		n, err := fmt.Scanln(&outputFile)
		if n != 1 {
			fmt.Println("输入参数数量不对")
			continue
		}
		if err != nil {
			fmt.Println(err)
			continue
		}
		break
	}
	writeExcel, err := excelize.OpenFile(outputFile)
	if err != nil {
		fmt.Println(err)
		return
	}

	//目标文件工作表名输入
	outSheet := ""
	for {
		fmt.Print("请输入要匹配的目标文件工作表名:")
		n, err := fmt.Scanln(&outSheet)
		if n != 1 {
			fmt.Println("输入参数数量不对")
			continue
		}
		if err != nil {
			fmt.Println(err)
			continue
		}
		break
	}
	writeRows := writeExcel.GetRows(outSheet)

	//目标文件要匹配的列输入
	outMatchCol := ""
	for {
		fmt.Print("请输入要匹配的源文件的列名:")
		n, err := fmt.Scanln(&outMatchCol)
		if n != 1 {
			fmt.Println("输入参数数量不对")
			continue
		}
		if err != nil {
			fmt.Println(err)
			continue
		}
		break
	}

	//目标文件要复制到的列输入
	outputCol := ""
	for {
		fmt.Print("请输入目标文件要复制到的列名:")
		n, err := fmt.Scanln(&outputCol)
		if n != 1 {
			fmt.Println("输入参数数量不对")
			continue
		}
		if err != nil {
			fmt.Println(err)
			continue
		}
		break
	}
	//outColNum := excelize.TitleToNumber(outputCol)

	countW := 1
	fmt.Println("outN", "\t", "inN", "\t", "匹配项", "\t\t", "复制内容")
	for _, row := range writeRows {
		countR := 1
		for _, rowR := range readRows {
			if rowR[excelize.TitleToNumber(inMatchCol)] == row[excelize.TitleToNumber(outMatchCol)] && rowR[excelize.TitleToNumber(inMatchCol)] != "" {
				fmt.Print(countW, "\t", countR, "\t", row[excelize.TitleToNumber(outMatchCol)], "\t")
				fmt.Println(rowR[inColNum])
				writeExcel.SetCellStr(outSheet, outputCol+strconv.Itoa(countW), rowR[inColNum])

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
