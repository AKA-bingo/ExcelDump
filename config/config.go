package config

import (
	"io/ioutil"
	"log"
	"os"

	"gopkg.in/yaml.v2"
)

type conf struct {
	DirLog       string `yaml:"DIR_LOG"`
	DirSheet     string `yaml:"DIR_SHEET"`
	SourceDir    string `yaml:"SOURCE_DIR"`
	ConventExt   string `yaml:"CONVENT_EXT"`
	OutPutExt    string `yaml:"OUTPUT_EXT"`
	PositionName string `yaml:"POSITION_NAME"`
}

var Conf = new(conf)

func init() {
	yamlFile, err := ioutil.ReadFile("config.yaml")
	if err != nil {
		log.Fatalf("Config file Get err #%v ", err)
	}
	err = yaml.Unmarshal(yamlFile, Conf)
	if err != nil {
		log.Fatalf("Unmarshal: %v", err)
	}

	//输出重定向
	f, _ := os.OpenFile(Conf.DirLog, os.O_WRONLY|os.O_CREATE|os.O_SYNC|os.O_APPEND, 0755)
	log.SetOutput(f)
}
