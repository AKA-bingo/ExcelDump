package main

import (
	"io/ioutil"
	"log"
	"os"
	"strings"

	"github.com/AKA-bingo/ExcelDump/config"
	"github.com/AKA-bingo/ExcelDump/service"
)

func main() {
	files, err := ioutil.ReadDir(config.Conf.SourceDir)
	if err != nil {
		log.Fatalf("获取源数据文件夹下文件失败, err:%v", err)
	}

	//文件分隔符
	//string(os.PathSeparator)

	for _, file := range files {
		if !file.IsDir() && strings.HasSuffix(file.Name(), config.Conf.ConventExt) {
			service.FileConvent(config.Conf.SourceDir + string(os.PathSeparator) + file.Name())
		}

	}

}
