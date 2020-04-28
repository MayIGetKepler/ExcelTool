package ExcelTool
import (
	"errors"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/go-flutter-desktop/go-flutter"
	"github.com/go-flutter-desktop/go-flutter/plugin"
)

const (
	ChannelName = "kepler.me.go.plugin.excelMerger"
	MethodMerge = "merge"
	KeyPaths    = "paths"
	KeyTarget   = "target"
	sheet       = "Sheet1"
)

type Plugin struct {
}

var _ flutter.Plugin = &Plugin{}

func (Plugin) InitPlugin(messenger plugin.BinaryMessenger) error {
	channel := plugin.NewMethodChannel(messenger, ChannelName, plugin.StandardMethodCodec{})
	channel.HandleFunc(MethodMerge, merge)
	return nil
}

func merge(arguments interface{}) (reply interface{}, err error) {
	args := arguments.(map[interface{}]interface{})
	paths := args[KeyPaths].([]string)
	targetPath := args[KeyTarget].(string)
	err = mergeInternal(paths, targetPath)
	return nil, err
}

func mergeInternal(paths []string, targetPath string) error {
	var errorMsg string
	newFile := excelize.NewFile()
	startIndex := 0
	for _, path := range paths {
		file, err := excelize.OpenFile(path)
		if err != nil {
			errorMsg += "open file : " + path + " error - " + err.Error() + "\n"
			continue
		}
		rows, err := file.GetRows(sheet)
		if err != nil {
			errorMsg += "file : " + path + "get rows error - " + err.Error() + "\n"
			continue
		}
		for i, row := range rows {
			copyRows := 0
			for k, v := range row {
				cellName, err := excelize.CoordinatesToCellName(k+1, startIndex+i+1)
				if err != nil {
					errorMsg += "file : " + path + "get cell name error - " + err.Error() + "\n"
					continue
				}
				err = newFile.SetCellValue(sheet, cellName, v)
				copyRows++
			}
			startIndex += copyRows
		}
	}
	err := newFile.SaveAs(targetPath)
	if err != nil {
		errorMsg += "save file error - " + err.Error() + "\n"
	}
	if errorMsg != "" {
		return errors.New(errorMsg)
	} else {
		return nil
	}
}

