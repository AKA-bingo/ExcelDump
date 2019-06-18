package main

import (
	"fmt"
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

	fmt.Println("Start Convent")
	for _, file := range files {
		if !file.IsDir() && strings.HasSuffix(file.Name(), config.Conf.ConventExt) && !strings.HasPrefix(file.Name(), "~$") {
			service.FileConvent(config.Conf.SourceDir + string(os.PathSeparator) + file.Name())
		}

	}
	fmt.Print("Convent done\nInput anything to exist program:")
	var pause string
	fmt.Scan(&pause)
}
