package internal

import (
	"bytes"
	"fmt"
	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"
	"log/slog"
	"sort"
	"strconv"
	"strings"
)

type DesignerMoney struct {
	Name  string
	Money float64
}

type DesignerMoneyList []DesignerMoney

func (dl DesignerMoneyList) Len() int {
	return len(dl)
}

func (dl DesignerMoneyList) Less(i, j int) bool {
	return dl[i].Money > dl[j].Money
}

func (dl DesignerMoneyList) Swap(i, j int) {
	dl[i], dl[j] = dl[j], dl[i]
}

func ReadExcelFileFromHttpMultipartFileHeader(c *gin.Context) (excelFile *excelize.File, oriFileName string, err error) {
	//获取文件
	_, fileUpHeader, err := c.Request.FormFile("file")
	if err != nil {
		slog.Error("从http form中获取文件失败", "err", err)
		return
	}

	file, err := fileUpHeader.Open()
	if err != nil {
		slog.Error("打开文件失败", "err", err)
		return
	}

	excelFile, err = excelize.OpenReader(file, excelize.Options{RawCellValue: true})
	if err != nil {
		slog.Error("打开excel文件失败", "err", err)
		return
	}

	oriFileName = fileUpHeader.Filename
	return
}

func CalExcel(excelFile *excelize.File, ignoreExcelLastRow bool, contractMoneyColumnName string, designerColumnNameList []string) (newExcelBytes *bytes.Buffer, err error) {

	defer func() {
		// Close the spreadsheet.
		if err1 := excelFile.Close(); err1 != nil {
			slog.Error("关闭excel文件失败", "err1", err1)
		}
	}()

	sheetList := excelFile.GetSheetList()
	if len(sheetList) != 1 {
		err = fmt.Errorf("excel文件缺少sheet")
		slog.Error("excel文件缺少sheet")
		return
	}

	sheetName := sheetList[0]
	rows, err := excelFile.GetRows(sheetName, excelize.Options{RawCellValue: true})
	if err != nil {
		slog.Error("获取excel文件的sheet失败", "err", err)
		return
	}

	if len(rows) < 1 {
		err = fmt.Errorf("excel文件的sheet缺少表头")
		slog.Error("excel文件的sheet缺少表头")
		return
	}

	// 处理表头
	var (
		designerColumnNameListIndex  = make([]int, 0, len(designerColumnNameList))
		contractMoneyColumnNameIndex = -1 // 合同金额列index
	)

	// 数据量不大,双循环
	for index, name := range rows[0] {
		// 设计师列名
		for _, columnName := range designerColumnNameList {
			if name == columnName {
				designerColumnNameListIndex = append(designerColumnNameListIndex, index)
			}
		}

		// 合同金额列名
		if name == contractMoneyColumnName {
			contractMoneyColumnNameIndex = index
		}
	}

	if len(designerColumnNameListIndex) != len(designerColumnNameList) {
		slog.Error("excel文件里的sheet设计师列名配置在excel中不存在")
		err = fmt.Errorf("excel文件里的sheet设计师列名配置在excel中不存在")
		return
	}

	if contractMoneyColumnNameIndex == -1 {
		slog.Error("excel文件里的sheet合同金额列名配置在excel中不存在")
		err = fmt.Errorf("excel文件里的sheet合同金额列名配置在excel中不存在")
		return
	}

	// 读取数据
	if len(rows) < 2 {
		slog.Error("excel文件的sheet缺少数据行")
		err = fmt.Errorf("excel文件的sheet缺少数据行")
		return
	}

	var (
		rowEndIndex      int
		designerMoneyMap = make(map[string]float64)
	)

	if ignoreExcelLastRow {
		rowEndIndex = len(rows) - 1
	} else {
		rowEndIndex = len(rows)
	}

	for i := 1; i < rowEndIndex; i++ {
		row := rows[i]
		// 合同金额
		if len(row) <= contractMoneyColumnNameIndex {
			msg := fmt.Sprintf("excel文件第%d行合同金额列数据不存在", i+1)
			slog.Error(msg)
			err = fmt.Errorf(msg)
			return
		}
		if row[contractMoneyColumnNameIndex] == "" {
			msg := fmt.Sprintf("excel文件第%d行合同金额列数据为空", i+1)
			slog.Error(msg)
			err = fmt.Errorf(msg)
			return
		}
		var money float64
		money, err = strconv.ParseFloat(row[contractMoneyColumnNameIndex], 64)
		if err != nil {
			msg := fmt.Sprintf("excel文件第%d行合同金额列数据格式错误 : %v\n", i+1, err)
			slog.Error(msg)
			err = fmt.Errorf(msg)
			return
		}

		// 设计师列表
		var designerList = make([]string, 0)
		for _, index := range designerColumnNameListIndex {
			if len(row) <= index {
				msg := fmt.Sprintf("excel文件第%d行'%s'列数据不存在", i+1, designerList[index])
				slog.Error(msg)
				err = fmt.Errorf(msg)
				return
			}
			if strings.Trim(row[index], " ") == "" {
				continue
			} else {
				designerList = append(designerList, row[index])
			}
		}

		if len(designerList) == 0 {
			msg := fmt.Sprintf("excel文件第%d行设计师数据为空", i+1)
			slog.Error(msg)
			err = fmt.Errorf(msg)
			return
		}

		// 计算每个设计师的金额
		moneyPerDesigner := money / float64(len(designerList))
		for _, designer := range designerList {
			if _, ok := designerMoneyMap[designer]; ok {
				designerMoneyMap[designer] += moneyPerDesigner
			} else {
				designerMoneyMap[designer] = moneyPerDesigner
			}
		}
	}

	// 转换成DesignerMoneyList
	designerMoneyData := transferMapDataToSortedDesignerMoneyList(designerMoneyMap)

	// 生成excel文件
	var (
		f                  *excelize.File
		generateSheetName  = "Sheet1"
		generateSheetIndex int
		defaultStyleIndex  int
	)

	f = excelize.NewFile()
	generateSheetIndex, err = f.NewSheet(generateSheetName)
	if err != nil {
		msg := fmt.Sprintf("创建excel文件失败: %v\n", err)
		slog.Error(msg)
		err = fmt.Errorf(msg)
		return
	}

	style := excelize.Style{
		Protection: &excelize.Protection{
			Locked: true,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
			WrapText:   true,
		},
	}
	defaultStyleIndex, err = f.NewStyle(&style)

	err = f.SetRowStyle(sheetName, 1, 1, defaultStyleIndex)
	// 设置表头
	err = f.SetCellStr(generateSheetName, "A1", "姓名")
	if err != nil {
		msg := fmt.Sprintf("设置excel文件第%d行第%d列失败: %v\n", 1, 1, err)
		slog.Error(msg)
		err = fmt.Errorf(msg)
		return
	}
	err = f.SetCellStr(generateSheetName, "B1", "金额")
	if err != nil {
		msg := fmt.Sprintf("设置excel文件第%d行第%d列失败: %v\n", 1, 2, err)
		slog.Error(msg)
		err = fmt.Errorf(msg)
		return
	}

	var totalMoney float64
	for i, item := range designerMoneyData {
		err = f.SetRowStyle(sheetName, i+1, i+1, defaultStyleIndex)
		totalMoney += item.Money
		err = f.SetCellStr(generateSheetName, fmt.Sprintf("A%d", i+2), item.Name)
		if err != nil {
			msg := fmt.Sprintf("设置excel文件第%d行第%d列失败: %v\n", i+2, 1, err)
			slog.Error(msg)
			err = fmt.Errorf(msg)
			return
		}
		err = f.SetCellFloat(generateSheetName, fmt.Sprintf("B%d", i+2), item.Money, 2, 64)
		if err != nil {
			msg := fmt.Sprintf("设置excel文件第%d行第%d列失败: %v\n", i+2, 2, err)
			slog.Error(msg)
			err = fmt.Errorf(msg)
			return
		}
	}

	// 记录总金额
	err = f.SetCellFloat(generateSheetName, fmt.Sprintf("B%d", len(designerMoneyData)+2), totalMoney, 2, 64)
	if err != nil {
		msg := fmt.Sprintf("设置excel文件第%d行第%d列失败: %v\n", len(designerMoneyData)+2, 2, err)
		slog.Error(msg)
		err = fmt.Errorf(msg)
		return
	}

	f.SetActiveSheet(generateSheetIndex)
	newExcelBytes, err = f.WriteToBuffer()
	if err != nil {
		msg := fmt.Sprintf("生成excel文件失败: %v\n", err)
		slog.Error(msg)
		err = fmt.Errorf(msg)
		return
	}
	slog.Info("生成excel文件完成")
	return
}

func transferMapDataToSortedDesignerMoneyList(oriData map[string]float64) (data DesignerMoneyList) {
	for name, money := range oriData {
		data = append(data, DesignerMoney{
			Name:  name,
			Money: money,
		})
	}

	sort.Sort(data)
	return
}
