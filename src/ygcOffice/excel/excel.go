package excel

import (
	"github.com/alecthomas/log4go"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"runtime"
	"strings"
)

type ExcelObject struct {
	File                 *excelize.File
	X, Y, Xs, Ys, Xe, Ye int
	Operation            string
	CurrentSheet         string
}

func GetCellName(columnIndex, rowIndex int) string {
	return excelize.ToAlphaString(columnIndex) + fmt.Sprintf("%d", rowIndex)
}

func FindConlumnCell(file *excelize.File, sheetName string, columnStart int, columnEnd int, row int, findVal string) (result bool, columnIndex int) {
	result, columnIndex, _ = FindCell(file, sheetName, columnStart, columnEnd, row, row, findVal)
	return result, columnIndex
}

func FindRowCell(file *excelize.File, sheetName string, column int, rowStart int, rowEnd int, findVal string) (result bool, rowIndex int) {
	result, _, rowIndex = FindCell(file, sheetName, column, column, rowStart, rowEnd, findVal)
	return result, rowIndex
}

//func FindAllCell(file *excelize.File, sheetName string, findVal string) (result bool, columnIndex, rowIndex int) {
//	return FindCell(file, sheetName, 0, 1000, 1, 30000, findVal)
//}

func FindStartTextCell(file *excelize.File, sheetName string,xstart,ystart int, findVal string) (result bool, columnIndex, rowIndex int) {
	return FindCell(file, sheetName, xstart, xstart+500, ystart, ystart+5000, findVal)
}

func FindCell(file *excelize.File, sheetName string, columnStart int, columnEnd int, rowStart int, rowEnd int, findVal string) (result bool, columnIndex, rowIndex int) {
	if rowEnd < rowStart {
		rowEnd = rowStart + 5000
	}

	if columnEnd < columnStart {
		columnEnd = columnStart + 500
	}

	var cellName string
	var tmp string
	for l := rowStart; l <= rowEnd; l++ {
		for i := columnStart; i <= columnEnd; i++ {
			cellName=GetCellName(i, l)
			//tmp1:= file.GetCellFormula(sheetName,cellName)
			tmp = file.GetCellValue(sheetName, cellName)
			//if strings.Index(tmp,"\n")>=0{
			//	tmp= strings.Replace(tmp,"\n","",-1)
			//}
			if strings.Index(tmp,findVal)==0{
				return true, i, l
			}
		}
	}
	log4go.Error("未找到结束标识：%s 在文件 %s Sheet %s", findVal,file.Path,sheetName)
	return false, columnEnd, rowEnd
}

func GetCompnyNameFromPath(path string) string {
	var splitPath []string
	if runtime.GOOS == "windows" {
		splitPath = strings.Split(path, "\\")
	} else {
		splitPath = strings.Split(path, "/")
	}

	splitFile := strings.Split(splitPath[len(splitPath)-1], ".")
	splitName := strings.Split(splitFile[0], "-")
	val := []byte(splitName[2])[:strings.Index(splitName[2], "20")]
	return string(val)
}
