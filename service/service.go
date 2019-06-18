package service

import (
	"errors"
	"fmt"
	"io/ioutil"
	"log"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize/v2"

	"github.com/AKA-bingo/ExcelDump/config"
)

var (
	ErrConventRuleNotFound = errors.New("ExcelDump Error Convent Rule Not Found")
	ErrTableNameNotFound   = errors.New("ExcelDump Error Table Name Not Found")
)

func FileConvent(filePath string) {
	file, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Printf("excelize.OpenFile(%v) err:%v", filePath, err)
		return
	}

	ConventRule, err := GetOutPutRule(file)
	if err != nil {
		log.Printf("GetOutPutRule(%v) err:%v", file.Path, err)
		return
	} else if len(ConventRule) <= 0 {
		log.Printf("GetOutPutRule err:%v", ErrConventRuleNotFound)
		return
	}

	sheets := file.GetSheetMap()
	for _, sheet := range sheets {
		if sheet == config.Conf.DirSheet {
			continue
		}
		err := ConventSheet(file, sheet, ConventRule)
		if err != nil {
			log.Printf("ConventSheet(%v, %v, %v) err:%v", file.Path, sheet, ConventRule[sheet], err)
		}
	}

	//fmt.Printf("%+v\n", ConventRule)
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
	//fmt.Printf("%+v", rows)

	return DirMap, nil
}

func ConventSheet(file *excelize.File, sheetName string, conventRule map[string]string) error {
	tableName, err := file.GetCellValue(sheetName, config.Conf.TableNamePos)
	if err != nil {
		log.Printf("GetCellValue(%v, %v) err%v", sheetName, config.Conf.TableNamePos, err)
		return err
	} else if tableName == "" {
		return ErrTableNameNotFound
	}

	destFilePath := conventRule[tableName] + string(os.PathSeparator) + tableName + config.Conf.OutPutExt
	d1 := []byte("hello\ngo\n")
	err = ioutil.WriteFile(destFilePath, d1, 0644)
	fmt.Println(err)
	return nil
}
