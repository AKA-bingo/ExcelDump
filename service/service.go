package service

import (
	"errors"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"

	"github.com/AKA-bingo/ExcelDump/config"
)

var (
	ErrConventRuleNotFound  = errors.New("ExcelDump Error Convent Rule Not Found")
	ErrTableNameNotFound    = errors.New("ExcelDump Error Table Name Not Found")
	ErrContentNotFound      = errors.New("ExcelDump Error Content Not Found")
	ErrPositionNameNotFound = errors.New("ExcelDump Error Position Name Not Found")
	ErrConventRouteNotFound = errors.New("ExcelDump Error Convent Route Not Found")
)

func FileConvent(filePath string) {
	file, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Printf("excelize.OpenFile(%v) err:%v", filePath, err)
		return
	}
	fmt.Printf("Start convent excel %v\n", filepath.Base(file.Path))

	ConventRule, err := GetOutPutRule(file)
	if err != nil {
		log.Printf("GetOutPutRule(%v) err:%v", file.Path, err)
		return
	} else if len(ConventRule) <= 0 {
		log.Printf("GetOutPutRule err:%v", ErrConventRuleNotFound)
		return
	}
	fmt.Printf("Get convent router:\n")
	for table, info := range ConventRule {
		fmt.Printf("%v => %v%v%v.txt\n", table, info["PATH"], string(os.PathSeparator), table)
	}

	sheets := file.GetSheetMap()
	for _, sheet := range sheets {
		if sheet == config.Conf.DirSheet {
			continue
		}
		err := ConventSheet(file, sheet, ConventRule)
		if err != nil {
			log.Printf("ConventSheet(%v, %v, ConventRule) err:%v", file.Path, sheet, err)
		}
	}

	fmt.Printf("Convent excel %v done\n", filepath.Base(file.Path))
}

//定位标识符所在位置
func Position(file *excelize.File, sheetName string) (int, int, error) {
	rows, err := file.GetRows(sheetName)
	if err != nil {
		return 0, 0, err
	}
	var rowPos, columnPos int

	for i, row := range rows {
		for j, value := range row {
			if value == config.Conf.PositionName {
				rowPos, columnPos = i, j
				return rowPos, columnPos, nil
			}
		}
	}

	return 0, 0, ErrPositionNameNotFound
}

func GetOutPutRule(file *excelize.File) (map[string]map[string]string, error) {
	rows, err := file.GetRows(config.Conf.DirSheet)
	if err != nil {
		return nil, err
	}

	rowPos, columnPos, err := Position(file, config.Conf.DirSheet) //3, 0

	if len(rows) <= rowPos+1 {
		return nil, ErrConventRouteNotFound
	}

	DirMap := make(map[string]map[string]string, len(rows)-rowPos-1)

	for _, row := range rows[rowPos+1:] {
		if row[columnPos+2] != "" && row[columnPos+3] != "" {
			dirInfo := make(map[string]string, 2)
			dirInfo["PATH"] = row[columnPos+3]
			dirInfo["INIT"] = "1" //init = 1 : 需要初始化（清空文件内容） inti = 0 : 追加文件内容（去除表头）
			DirMap[row[columnPos+2]] = dirInfo
		}
	}

	return DirMap, nil
}

func ConventSheet(file *excelize.File, sheetName string, conventRule map[string]map[string]string) error {
	column_len := 0

	rowPos, columnPos, err := Position(file, sheetName)
	if err != nil {
		log.Printf("Position(%v, %v) err:%v", filepath.Base(file.Path), sheetName, err)
		return err
	}

	contents, err := file.GetRows(sheetName)
	if err != nil {
		log.Printf("file.GetRows(%v) err:%v", sheetName, err)
		return err
	} else if len(contents) <= rowPos+1 {
		return ErrContentNotFound
	}

	tableName := contents[rowPos+1][columnPos]
	if tableName == "" {
		return ErrTableNameNotFound
	}

	fmt.Printf("Start convent table %v ... ", tableName)

	//创建目录
	if _, err := os.Stat(conventRule[tableName]["PATH"]); err != nil {
		err = os.MkdirAll(conventRule[tableName]["PATH"], 0755)
		if err != nil {
			log.Printf("os.MkdirAll(%v, 0755) err:%v", conventRule[tableName], err)
			return err
		}
	}

	destFilePath := conventRule[tableName]["PATH"] + string(os.PathSeparator) + tableName + config.Conf.OutPutExt
	flag := os.O_APPEND | os.O_CREATE | os.O_WRONLY
	if conventRule[tableName]["INIT"] == "1" {
		flag = flag | os.O_TRUNC
		conventRule[tableName]["INIT"] = "0"
	} else {
		column_len, _ = strconv.Atoi(conventRule[tableName]["COLUMN_LEN"])
		rowPos += 2
	}

	txtFile, err := os.OpenFile(destFilePath, flag, 0755)
	if err != nil {
		log.Printf("os.OpenFile(%v, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0600) err:%v", destFilePath, err)
		return err
	}
	defer txtFile.Close()

	for _, row := range contents[rowPos:] {
		content := make([]string, 0)
		for i, column := range row[columnPos+1:] {
			if column == "" && column_len == 0 {
				column_len = i
				break
			} else if i > column_len && column_len != 0 {
				break
			}
			content = append(content, column)
		}
		if column_len == 0 {
			column_len = len(row[columnPos+1:])
		}
		if _, err = txtFile.WriteString(strings.Join(content, "	") + "\n"); err != nil {
			log.Printf("txtFile.WriteString(%v + \n) err:%v", strings.Join(content, "	"), err)
		}
	}

	conventRule[tableName]["COLUMN_LEN"] = strconv.Itoa(column_len)

	fmt.Printf("done\n")

	return nil
}
