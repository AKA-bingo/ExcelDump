package config

import (
	"io/ioutil"
	"log"

	"gopkg.in/yaml.v2"
)

type conf struct {
	DirSheet           string `yaml:"DIR_SHEET"`
	SourceDir          string `yaml:"SOURCE_DIR"`
	ConventExt         string `yaml:"CONVENT_EXT"`
	OutPutExt          string `yaml:"OUTPUT_EXT"`
	TableNamePos       string `yaml:"TABLE_NAME_POS"`
	ContentStartRow    string `yaml:"CONTENT_START_ROW"`
	ContentStartColumn string `yaml:"CONTENT_START_COLUMN"`
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
}
