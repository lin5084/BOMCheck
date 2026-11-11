package main

import (
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"

	"github.com/xuri/excelize/v2"
	"golang.org/x/sys/windows"
)

func init() {
	// Windows设置控制台UTF‑8，解决CMD乱码
	stdin := windows.Handle(os.Stdin.Fd())
	var mode uint32
	_ = windows.GetConsoleMode(stdin, &mode)
	_ = windows.SetConsoleMode(stdin, mode|windows.ENABLE_VIRTUAL_TERMINAL_PROCESSING)
	_ = windows.SetConsoleOutputCP(65001)
}

type MergeInfo struct {
	startRow int
	endRow   int
	startCol int
	endCol   int
}

type CheckResult struct {
	fileName      string
	countMismatch []string // "第X行"
	mergeLeak     []string // "第X行"
	dupRefDes     map[string][]int // key:位号 value:行号数组
	dupPartNo     map[string][]int // key:料号 value:行号数组
}

// 分割位号：中文逗号，英文逗号，文字“和”
func splitRefDes(s string) []string {
	s = strings.TrimSpace(s)
	if s == "" {
		return nil
	}
	re := regexp.MustCompile(`[,，和]`)
	parts := re.Split(s, -1)
	var res []string
	for _, p := range parts {
		p = strings.TrimSpace(p)
		if p != "" {
			res = append(res, p)
		}
	}
	return res
}

// 获取合并单元格完整信息，返回每个合并块；再建立map：row-col → *MergeInfo
func getMergeMap(f *excelize.File, sheet string) (map[string]*MergeInfo, []MergeInfo, error) {
	mergeCellMap := make(map[string]*MergeInfo)
	var mergeList []MergeInfo
	ranges, err := f.GetMergeCells(sheet)
	if err != nil {
		return nil, nil, err
	}
	for _, rng := range ranges {
		sAxis, eAxis := rng.GetStartAxis(), rng.GetEndAxis()
		colS, rowS, _ := excelize.CellNameToCoordinates(sAxis)
		colE, rowE, _ := excelize.CellNameToCoordinates(eAxis)
		mi := &MergeInfo{
			startRow: rowS,
			endRow:   rowE,
			startCol: colS,
			endCol:   colE,
		}
		mergeList = append(mergeList, *mi)
		for r := rowS; r <= rowE; r++ {
			for c := colS; c <= colE; c++ {
				key := fmt.Sprintf("%d-%d", r, c)
				mergeCellMap[key] = mi
			}
		}
	}
	return mergeCellMap, mergeList, nil
}

// 获取单元格值：如果属于合并单元格，取合并块的首行首列的值
func getMergedCellValue(f *excelize.File, sheet string, row int, col int, mergeCellMap map[string]*MergeInfo) (string, error) {
	key := fmt.Sprintf("%d-%d", row, col)
	mi, ok := mergeCellMap[key]
	cellAxis, _ := excelize.CoordinatesToCellName(col, row)
	if !ok {
		// 非合并单元格，直接取值
		return f.GetCellValue(sheet, cellAxis)
	}
	// 是合并单元格，取合并起始单元格
	firstAxis, _ := excelize.CoordinatesToCellName(mi.startCol, mi.startRow)
	return f.GetCellValue(sheet, firstAxis)
}

func processFile(path string) (*CheckResult, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	sheet := f.GetSheetList()[0]
	// A1单元格必须等于“料号”
	a1, err := f.GetCellValue(sheet, "A1")
	if err != nil {
		return nil, err
	}
	if a1 != "料号" {
		return nil, fmt.Errorf("not bom")
	}

	mergeCellMap, _, err := getMergeMap(f, sheet)
	if err != nil {
		return nil, err
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}
	res := &CheckResult{
		fileName:  filepath.Base(path),
		dupRefDes: make(map[string][]int),
		dupPartNo: make(map[string][]int),
	}

	partNoMap := make(map[string][]int)
	refDesAll := make(map[string][]int)

	maxRow := len(rows)
	// E=5列，F=6列（excel列号，A=1）
	const colE = 5
	const colF = 6

	// 从第4行开始，数组下标3对应Excel第4行
	for rowIdx := 3; rowIdx < maxRow; rowIdx++ {
		excelRow := rowIdx + 1

		// A列直接拿原始行数据
		row := rows[rowIdx]
		aVal := ""
		if 0 < len(row) {
			aVal = strings.TrimSpace(row[0])
		}

		// E、F列：如果是合并单元格，自动取合并块首行的值
		eVal, errE := getMergedCellValue(f, sheet, excelRow, colE, mergeCellMap)
		if errE != nil {
			eVal = ""
		}
		eVal = strings.TrimSpace(eVal)

		fVal, errF := getMergedCellValue(f, sheet, excelRow, colF, mergeCellMap)
		if errF != nil {
			fVal = ""
		}
		fVal = strings.TrimSpace(fVal)

		// 料号重复检查
		if aVal != "" {
			partNoMap[aVal] = append(partNoMap[aVal], excelRow)
		}

		refList := splitRefDes(fVal)
		var eNum int
		_, _ = fmt.Sscanf(eVal, "%d", &eNum)

		// 判断：本行是合并块的非首行，则跳过【组件数量 vs 位号数量校验】
		isMergeSubRow := false
		keyE := fmt.Sprintf("%d-%d", excelRow, colE)
		if mi, ok := mergeCellMap[keyE]; ok && mi.startRow != excelRow {
			isMergeSubRow = true
		}

		// 只有合并首行 / 非合并单元格才做数量比对；合并内部子行跳过，不再误报
		if !isMergeSubRow && eVal != "" {
			realCnt := len(refList)
			if realCnt != eNum {
				res.countMismatch = append(res.countMismatch, fmt.Sprintf("%d", excelRow))
			}
		}

		// 收集位号重复
		for _, rd := range refList {
			refDesAll[rd] = append(refDesAll[rd], excelRow)
		}

		// ============修复漏合并检测逻辑============
		if rowIdx+1 < maxRow {
			nextRowIdx := rowIdx + 1
			nextExcelRow := nextRowIdx + 1

			nextRow := rows[nextRowIdx]
			nextA := ""
			if len(nextRow) > 0 {
				nextA = strings.TrimSpace(nextRow[0])
			}

			nextEVal, errNE := getMergedCellValue(f, sheet, nextExcelRow, colE, mergeCellMap)
			if errNE != nil {
				nextEVal = ""
			}
			nextEVal = strings.TrimSpace(nextEVal)

			nextFVal, errNF := getMergedCellValue(f, sheet, nextExcelRow, colF, mergeCellMap)
			if errNF != nil {
				nextFVal = ""
			}
			nextFVal = strings.TrimSpace(nextFVal)

			keyNextE := fmt.Sprintf("%d-%d", nextExcelRow, colE)
			keyNextF := fmt.Sprintf("%d-%d", nextExcelRow, colF)
			nextEIsMerged := false
			if _, ok := mergeCellMap[keyNextE]; ok {
				nextEIsMerged = true
			}
			nextFIsMerged := false
			if _, ok := mergeCellMap[keyNextF]; ok {
				nextFIsMerged = true
			}

			curEFHasData := eVal != "" || fVal != ""
			nextAHasData := nextA != ""
			nextEFAllEmpty := (nextEVal == "") && (nextFVal == "")
			nextEFNotCover := (!nextEIsMerged) || (!nextFIsMerged)

			if curEFHasData && nextAHasData && nextEFAllEmpty && nextEFNotCover {
				res.mergeLeak = append(res.mergeLeak, fmt.Sprintf("%d", excelRow))
			}
		}
	}

	// 重复料号输出
	for pn, rowsArr := range partNoMap {
		if len(rowsArr) > 1 {
			res.dupPartNo[pn] = rowsArr
		}
	}
	// 重复位号输出
	for rd, rowsArr := range refDesAll {
		if len(rowsArr) > 1 {
			res.dupRefDes[rd] = rowsArr
		}
	}

	return res, nil
}

func main() {
	workDir := "."
	if len(os.Args) > 1 {
		workDir = os.Args[1]
	}

	fmt.Printf("BOM 检查目录：%s\n\n", workDir)

	var bomFiles []string
	_ = filepath.WalkDir(workDir, func(path string, d fs.DirEntry, err error) error {
		if err != nil {
			return nil
		}
		if d.IsDir() {
			return nil
		}
		base := filepath.Base(path)
		if strings.HasPrefix(base, "~$") {
			return nil
		}
		if strings.ToLower(filepath.Ext(path)) == ".xlsx" {
			bomFiles = append(bomFiles, path)
		}
		return nil
	})

	var results []*CheckResult
	for _, fpath := range bomFiles {
		r, err := processFile(fpath)
		if err != nil {
			continue
		}
		results = append(results, r)
	}

	if len(results) == 0 {
		fmt.Println("当前文件夹下无BOM格式文件。")
		fmt.Println("按回车键退出...")
		_, _ = os.Stdin.Read(make([]byte, 1))
		return
	}

	totalErr := 0
	for _, res := range results {
		// 组件数量不一致
		for _, rowStr := range res.countMismatch {
			fmt.Printf("%s 中的第 %s 行组件数量和实际位号数量不一致。\n", res.fileName, rowStr)
			totalErr++
		}
		// 漏合并
		for _, rowStr := range res.mergeLeak {
			fmt.Printf("%s 中的第 %s 行存在单元格漏合并的情况。\n", res.fileName, rowStr)
			totalErr++
		}
		// 重复位号
		var dupRefList []string
		for k := range res.dupRefDes {
			dupRefList = append(dupRefList, k)
		}
		sort.Strings(dupRefList)
		for _, rd := range dupRefList {
			rows := res.dupRefDes[rd]
			var strs []string
			for _, v := range rows {
				strs = append(strs, fmt.Sprintf("%d", v))
			}
			fmt.Printf("%s 中的位号 %s 出现重复（出现在第 %s 行）。\n", res.fileName, rd, strings.Join(strs, "、"))
			totalErr++
		}
		// 重复料号
		var dupPartList []string
		for k := range res.dupPartNo {
			dupPartList = append(dupPartList, k)
		}
		sort.Strings(dupPartList)
		for _, pn := range dupPartList {
			rows := res.dupPartNo[pn]
			var strs []string
			for _, v := range rows {
				strs = append(strs, fmt.Sprintf("%d", v))
			}
			fmt.Printf("%s 中的料号 %s 出现重复（出现在第 %s 行）。\n", res.fileName, pn, strings.Join(strs, "、"))
			totalErr++
		}
	}

	fmt.Printf("\n检测完成，共扫描 %d 个BOM文件，发现 %d 个告警。\n", len(results), totalErr)
	fmt.Println("按回车键退出...")
	_, _ = os.Stdin.Read(make([]byte, 1))
}
