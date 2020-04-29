package ExcelTool

import (
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/go-flutter-desktop/go-flutter"
	"github.com/go-flutter-desktop/go-flutter/plugin"
	"strings"
)

const (
	ChannelName = "kepler.me.go.plugin.excelMerger"
	MethodMerge = "merge"
	KeyPaths    = "paths"
	KeyTarget   = "target"
	sheet       = "Sheet1"
)

var header = []string{"序号", "指示灯", "倒计时", "任务名称", "责任人", "审核人", "任务类型", "状态", "参与人", "完成率%", "计划开始日期", "计划完成日期", "实际完成日期", "估计工作量", "填报工作量", "确认工作量", "创建日期"}

type Plugin struct {
}

var _ flutter.Plugin = &Plugin{}

func (Plugin) InitPlugin(messenger plugin.BinaryMessenger) error {
	channel := plugin.NewMethodChannel(messenger, ChannelName, plugin.StandardMethodCodec{})
	channel.HandleFunc(MethodMerge, merge)
	return nil
}

func merge(arguments interface{}) (reply interface{}, err error) {
	args := arguments.([]interface{})
	arg := args[0].(map[interface{}]interface{})
	originalPathsStr := arg[KeyPaths].(string)

	pathsMapStr := originalPathsStr[1 : len(originalPathsStr)-2]
	paths := strings.Split(pathsMapStr, ",")
	for i, path := range paths {
		paths[i] = fixPath(path)
	}
	targetPath := arg[KeyTarget].(string)
	err = mergeInternal(paths, targetPath)
	reply = nil
	if err != nil {
		reply = err.Error()
	}
	return reply, nil
}
func fixPath(str string) string {
	str = strings.Trim(str, " ")
	str = strings.TrimLeft(str, "\"")
	str = strings.TrimRight(str, "\"")
	return str
}

func mergeInternal(paths []string, targetPath string) error {
	var errorMsg string
	skipLineCount := 0
	newFile := excelize.NewFile()
	err := addHeader(newFile)
	if err != nil {
		errorMsg += "add header error :" + err.Error() + "\n"
		fmt.Println(err)
	} else {
		skipLineCount--
	}
	startIndex := 0
	for _, path := range paths {
		if path == "" {
			continue
		}
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
		copyRows := 0
		for i, row := range rows {
			if isHeader(row) {
				skipLineCount++
				fmt.Println("skip header")
				errorMsg += "skip header \n"
				copyRows++
				continue
			}
			if isBlankLine(row) {
				skipLineCount++
				errorMsg += "skip header \n"
				fmt.Println("skip blank")
				copyRows++
				continue
			}
			for k, v := range row {
				cellName, err := excelize.CoordinatesToCellName(k+1, startIndex+i+1-skipLineCount)
				if err != nil {
					errorMsg += "file : " + path + "get cell name error - " + err.Error() + "\n"
					continue
				}
				err = newFile.SetCellValue(sheet, cellName, v)
			}
			copyRows++
		}
		startIndex += copyRows
	}
	err = newFile.SaveAs(targetPath)
	if err != nil {
		errorMsg += "save file error - " + err.Error() + "\n"
	}
	if errorMsg != "" {
		return errors.New(errorMsg)
	} else {
		return nil
	}
}

func addHeader(file *excelize.File) error {
	return file.SetSheetRow(sheet, "A1", &header)
}

func isBlankLine(row []string) bool {
	if row == nil || len(row) == 0 {
		return true
	}
	for _, e := range row {
		if len(strings.Trim(e, " ")) > 0 {
			return false
		}
	}
	return true
}

func isHeader(row []string) bool {
	return StringSliceEqual(row, header)
}
func StringSliceEqual(a, b []string) bool {
	if len(a) != len(b) {
		return false
	}

	if (a == nil) != (b == nil) {
		return false
	}

	for i, v := range a {
		if v != b[i] {
			return false
		}
	}

	return true
}
