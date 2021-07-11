package main

import (
	"fmt"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func fillup(merge excelize.MergeCell, f *excelize.File, sheet string) {
	var mergeRange []string
	mergeCellRange := merge[0]
	fillValue := merge[1]
	mergeRange = strings.Split(mergeCellRange, ":")
	mergeRangeStart := mergeRange[0]
	mergeRangeEnd := mergeRange[1]
	_, rowNumberStart, err := excelize.SplitCellName(mergeRangeStart)
	if err != nil {
		fmt.Println(err)
		return
	}

	colName, rowNumberEnd, err := excelize.SplitCellName(mergeRangeEnd)
	if err != nil {
		fmt.Println(err)
		return
	}
	// fmt.Println(fillValue)
	// fmt.Println(colName)
	// fmt.Println(rowNumberStart)
	// fmt.Println(rowNumberEnd)
	// 取消合并单元格
	f.UnmergeCell(sheet, mergeRangeStart, mergeRangeEnd)
	// 填充单元格
	for i := rowNumberStart; i <= rowNumberEnd; i++ {
		cellName, err := excelize.JoinCellName(colName, i)
		if err != nil {
			fmt.Println(err)
		}
		// fmt.Println(cellName)
		err = f.SetCellValue(sheet, cellName, fillValue)
		if err != nil {
			fmt.Println(err)
			return
		}
		f.Save()
	}
}

func main() {
	f, err := excelize.OpenFile("test.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	sheets := f.GetSheetList()
	fmt.Println(sheets)
	for _, sheet := range sheets {
		// 获取所有合并的单元格
		mergeCells, _ := f.GetMergeCells(sheet)
		// fmt.Println(mergeCells)
		for _, mergeCell := range mergeCells {
			// 填充单元格
			fillup(mergeCell, f, sheet)
		}

		rows, err := f.GetRows(sheet)
		if err != nil {
			fmt.Println(err)
			return
		}
		header := rows[0]
		fmt.Println(header)
		for _, row := range rows[1:] {
			for _, colCell := range row {
				fmt.Print(colCell, "\t\t")
			}
			fmt.Println()
		}
	}

}
