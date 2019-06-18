package service

import (
	"errors"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"

	"github.com/AKA-bingo/ExcelDump/config"
)

var (
	ErrConventRuleNotFound = errors.New("ExcelDump Error Convent Rule Not Found")
	ErrTableNameNotFound   = errors.New("ExcelDump Error Table Name Not Found")
	ErrContentNotFound     = errors.New("ExcelDump Error Content Not Found")
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
	for table, path := range ConventRule {
		fmt.Printf("%v => %v%v%v.txt\n", table, path, string(os.PathSeparator), table)
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

//从第5行开始,提取第3,4列的数据
func GetOutPutRule(file *excelize.File) (map[string]string, error) {
	rows, err := file.GetRows(config.Conf.DirSheet)
	if err != nil {
		return nil, err
	}
	if len(rows) < 5 {
		return nil, nil
	}

	DirMap := make(map[string]string, len(rows)-4)

	for _, row := range rows[4:] {
		if row[2] != "" && row[3] != "" {
			DirMap[row[2]] = row[3]
		}
	}

	return DirMap, nil
}

func ConventSheet(file *excelize.File, sheetName string, conventRule map[string]string) error {
	tableName, err := file.GetCellValue(sheetName, config.Conf.TableNamePos)
	if err != nil {
		log.Printf("file.GetCellValue(%v, %v) err%v", sheetName, config.Conf.TableNamePos, err)
		return err
	} else if tableName == "" {
		return ErrTableNameNotFound
	}

	fmt.Printf("Start convent table %v ... ", tableName)

	//创建目录
	if _, err := os.Stat(conventRule[tableName]); err != nil {
		err = os.MkdirAll(conventRule[tableName], 0755)
		if err != nil {
			log.Printf("os.MkdirAll(%v, 0755) err:%v", conventRule[tableName], err)
			return err
		}
	}

	destFilePath := conventRule[tableName] + string(os.PathSeparator) + tableName + config.Conf.OutPutExt
	txtFile, err := os.OpenFile(destFilePath, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0755)
	if err != nil {
		log.Printf("os.OpenFile(%v, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0600) err:%v", destFilePath, err)
		return err
	}
	defer txtFile.Close()

	contents, err := file.GetRows(sheetName)
	if err != nil {
		log.Printf("file.GetRows(%v) err:%v", sheetName, err)
		return err
	} else if len(contents) < config.Conf.ContentStartRow {
		return ErrContentNotFound
	}

	column_len := 0
	for _, row := range contents[config.Conf.ContentStartRow-1:] {
		content := make([]string, 0)
		for i, column := range row[config.Conf.ContentStartColumn-1:] {
			if column == "" && column_len == 0 {
				column_len = i
				break
			} else if i > column_len && column_len != 0 {
				break
			}
			content = append(content, column)
		}
		if column_len == 0 {
			column_len = len(row[config.Conf.ContentStartColumn-1:]) - 1
		}
		if _, err = txtFile.WriteString(strings.Join(content, "	") + "\n"); err != nil {
			log.Printf("txtFile.WriteString(%v + \n) err:%v", strings.Join(content, "	"), err)
		}
	}

	fmt.Printf("done\n")

	return nil
}
