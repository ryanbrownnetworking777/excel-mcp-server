package excel

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

type ExcelizeExcel struct {
	file *excelize.File
}

func NewExcelizeExcel(file *excelize.File) Excel {
	return &ExcelizeExcel{file: file}
}

func (e *ExcelizeExcel) GetBackendName() string {
	return "excelize"
}

func (e *ExcelizeExcel) FindSheet(sheetName string) (Worksheet, error) {
	index, err := e.file.GetSheetIndex(sheetName)
	if err != nil {
		return nil, fmt.Errorf("sheet not found: %s", sheetName)
	}
	if index < 0 {
		return nil, fmt.Errorf("sheet not found: %s", sheetName)
	}
	return &ExcelizeWorksheet{file: e.file, sheetName: sheetName}, nil
}

func (e *ExcelizeExcel) CreateNewSheet(sheetName string) error {
	_, err := e.file.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("failed to create new sheet: %w", err)
	}
	return nil
}

func (e *ExcelizeExcel) CopySheet(srcSheetName string, destSheetName string) error {
	srcIndex, err := e.file.GetSheetIndex(srcSheetName)
	if srcIndex < 0 {
		return fmt.Errorf("source sheet not found: %s", srcSheetName)
	}
	if err != nil {
		return err
	}
	destIndex, err := e.file.NewSheet(destSheetName)
	if err != nil {
		return fmt.Errorf("failed to create destination sheet: %w", err)
	}
	if err := e.file.CopySheet(srcIndex, destIndex); err != nil {
		return fmt.Errorf("failed to copy sheet: %w", err)
	}
	srcNext := e.file.GetSheetList()[srcIndex+1]
	if srcNext != srcSheetName {
		e.file.MoveSheet(destSheetName, srcNext)
	}
	return nil
}

func (e *ExcelizeExcel) GetSheets() ([]Worksheet, error) {
	sheetList := e.file.GetSheetList()
	worksheets := make([]Worksheet, len(sheetList))
	for i, sheetName := range sheetList {
		worksheets[i] = &ExcelizeWorksheet{file: e.file, sheetName: sheetName}
	}
	return worksheets, nil
}

// SaveExcelize saves the Excel file to the specified path.
// Excelize's Save method restricts the file path length to 207 characters,
// but since this limitation has been relaxed in some environments,
// we ignore this restriction.
// https://github.com/qax-os/excelize/blob/v2.9.0/file.go#L71-L73
func (w *ExcelizeExcel) Save() error {
	file, err := os.OpenFile(filepath.Clean(w.file.Path), os.O_WRONLY|os.O_TRUNC|os.O_CREATE, os.ModePerm)
	if err != nil {
		return err
	}
	defer file.Close()
	return w.file.Write(file)
}

// FormatCells applies formatting to a range of cells using Excelize
// Excelizeを使ってセルの範囲にフォーマットを適用するのです！ 🎨(ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
func (e *ExcelizeExcel) FormatCells(sheetName, rangeStr string, style map[string]interface{}) error {
	// Create a new style with the provided formatting options
	// 提供されたフォーマットオプションで新しいスタイルを作成するのです！ ✨
	styleID, err := e.file.NewStyle(&excelize.Style{
		Font:      createFontStyle(style),
		Fill:      createFillStyle(style),
		Alignment: createAlignmentStyle(style),
		Border:    createBorderStyle(style),
		NumFmt:    getNumberFormat(style),
	})
	if err != nil {
		return fmt.Errorf("failed to create style: %w", err)
	}

	// Apply the style to the specified range
	// 指定された範囲にスタイルを適用するのです！ (◕‿◕)♡
	
	// Parse the range to get start and end cells
	// 範囲を解析して開始セルと終了セルを取得するのです！ ✨
	startCell, endCell, err := parseRangeForExcelize(rangeStr)
	if err != nil {
		return fmt.Errorf("failed to parse range %s: %w", rangeStr, err)
	}
	
	err = e.file.SetCellStyle(sheetName, startCell, endCell, styleID)
	if err != nil {
		return fmt.Errorf("failed to apply style to range %s: %w", rangeStr, err)
	}

	return nil
}

type ExcelizeWorksheet struct {
	file      *excelize.File
	sheetName string
}

func (w *ExcelizeWorksheet) Release() {
	// No resources to release in excelize
}

func (w *ExcelizeWorksheet) Name() (string, error) {
	return w.sheetName, nil
}

func (w *ExcelizeWorksheet) GetTables() ([]Table, error) {
	tables, err := w.file.GetTables(w.sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get tables: %w", err)
	}
	tableList := make([]Table, len(tables))
	for i, table := range tables {
		tableList[i] = Table{
			Name:  table.Name,
			Range: NormalizeRange(table.Range),
		}
	}
	return tableList, nil
}

func (w *ExcelizeWorksheet) GetPivotTables() ([]PivotTable, error) {
	pivotTables, err := w.file.GetPivotTables(w.sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get pivot tables: %w", err)
	}
	pivotTableList := make([]PivotTable, len(pivotTables))
	for i, pivotTable := range pivotTables {
		pivotTableList[i] = PivotTable{
			Name:  pivotTable.Name,
			Range: NormalizeRange(pivotTable.PivotTableRange),
		}
	}
	return pivotTableList, nil
}

func (w *ExcelizeWorksheet) SetValue(cell string, value any) error {
	if err := w.file.SetCellValue(w.sheetName, cell, value); err != nil {
		return err
	}
	if err := w.updateDimension(cell); err != nil {
		return fmt.Errorf("failed to update dimension: %w", err)
	}
	return nil
}

func (w *ExcelizeWorksheet) SetFormula(cell string, formula string) error {
	if err := w.file.SetCellFormula(w.sheetName, cell, formula); err != nil {
		return err
	}
	if err := w.updateDimension(cell); err != nil {
		return fmt.Errorf("failed to update dimension: %w", err)
	}
	return nil
}

func (w *ExcelizeWorksheet) GetValue(cell string) (string, error) {
	value, err := w.file.GetCellValue(w.sheetName, cell)
	if err != nil {
		return "", err
	}
	if value == "" {
		// try to get calculated value
		formula, err := w.file.GetCellFormula(w.sheetName, cell)
		if err != nil {
			return "", fmt.Errorf("failed to get formula: %w", err)
		}
		if formula != "" {
			return w.file.CalcCellValue(w.sheetName, cell)
		}
	}
	return value, nil
}

func (w *ExcelizeWorksheet) GetFormula(cell string) (string, error) {
	formula, err := w.file.GetCellFormula(w.sheetName, cell)
	if err != nil {
		return "", fmt.Errorf("failed to get formula: %w", err)
	}
	if formula == "" {
		// fallback
		return w.GetValue(cell)
	}
	if !strings.HasPrefix(formula, "=") {
		formula = "=" + formula
	}
	return formula, nil
}

func (w *ExcelizeWorksheet) GetDimention() (string, error) {
	return w.file.GetSheetDimension(w.sheetName)
}

func (w *ExcelizeWorksheet) GetPagingStrategy(pageSize int) (PagingStrategy, error) {
	return NewExcelizeFixedSizePagingStrategy(pageSize, w)
}

func (w *ExcelizeWorksheet) CapturePicture(captureRange string) (string, error) {
	return "", fmt.Errorf("CapturePicture is not supported in Excelize")
}

func (w *ExcelizeWorksheet) AddTable(tableRange, tableName string) error {
	enable := true
	if err := w.file.AddTable(w.sheetName, &excelize.Table{
		Range:             tableRange,
		Name:              tableName,
		StyleName:         "TableStyleMedium2",
		ShowColumnStripes: true,
		ShowFirstColumn:   false,
		ShowHeaderRow:     &enable,
		ShowLastColumn:    false,
		ShowRowStripes:    &enable,
	}); err != nil {
		return err
	}
	return nil
}

func (w *ExcelizeWorksheet) GetCellStyle(cell string) (*CellStyle, error) {
	styleID, err := w.file.GetCellStyle(w.sheetName, cell)
	if err != nil {
		return nil, fmt.Errorf("failed to get cell style: %w", err)
	}

	style, err := w.file.GetStyle(styleID)
	if err != nil {
		return nil, fmt.Errorf("failed to get style details: %w", err)
	}

	return convertExcelizeStyleToCellStyle(style), nil
}

func (w *ExcelizeWorksheet) SetCellStyle(cell string, style *CellStyle) error {
	excelizeStyle := convertCellStyleToExcelizeStyle(style)

	styleID, err := w.file.NewStyle(excelizeStyle)
	if err != nil {
		return fmt.Errorf("failed to create style: %w", err)
	}

	if err := w.file.SetCellStyle(w.sheetName, cell, cell, styleID); err != nil {
		return fmt.Errorf("failed to set cell style: %w", err)
	}

	return nil
}

func convertCellStyleToExcelizeStyle(style *CellStyle) *excelize.Style {
	result := &excelize.Style{}

	// Border
	if len(style.Border) > 0 {
		borders := make([]excelize.Border, len(style.Border))
		for i, border := range style.Border {
			excelizeBorder := excelize.Border{
				Type: border.Type.String(),
			}
			if border.Color != "" {
				excelizeBorder.Color = strings.TrimPrefix(border.Color, "#")
			}
			excelizeBorder.Style = borderStyleNameToInt(border.Style)
			borders[i] = excelizeBorder
		}
		result.Border = borders
	}

	// Font
	if style.Font != nil {
		font := &excelize.Font{}
		if style.Font.Bold != nil {
			font.Bold = *style.Font.Bold
		}
		if style.Font.Italic != nil {
			font.Italic = *style.Font.Italic
		}
		if style.Font.Underline != nil {
			font.Underline = style.Font.Underline.String()
		}
		if style.Font.Size != nil && *style.Font.Size > 0 {
			font.Size = float64(*style.Font.Size)
		}
		if style.Font.Strike != nil {
			font.Strike = *style.Font.Strike
		}
		if style.Font.Color != nil && *style.Font.Color != "" {
			font.Color = strings.TrimPrefix(*style.Font.Color, "#")
		}
		if style.Font.VertAlign != nil {
			font.VertAlign = style.Font.VertAlign.String()
		}
		result.Font = font
	}

	// Fill
	if style.Fill != nil {
		fill := excelize.Fill{}
		if style.Fill.Type != "" {
			fill.Type = style.Fill.Type.String()
		}
		fill.Pattern = fillPatternNameToInt(style.Fill.Pattern)
		if len(style.Fill.Color) > 0 {
			colors := make([]string, len(style.Fill.Color))
			for i, color := range style.Fill.Color {
				colors[i] = strings.TrimPrefix(color, "#")
			}
			fill.Color = colors
		}
		if style.Fill.Shading != nil {
			fill.Shading = fillShadingNameToInt(*style.Fill.Shading)
		}
		result.Fill = fill
	}

	// NumFmt
	if style.NumFmt != nil && *style.NumFmt != "" {
		result.CustomNumFmt = style.NumFmt
	}

	// DecimalPlaces
	if style.DecimalPlaces != nil && *style.DecimalPlaces > 0 {
		result.DecimalPlaces = style.DecimalPlaces
	}

	return result
}

func convertExcelizeStyleToCellStyle(style *excelize.Style) *CellStyle {
	result := &CellStyle{}

	// Border
	if len(style.Border) > 0 {
		var borders []Border
		for _, border := range style.Border {
			borderStyle := Border{
				Type: BorderType(border.Type),
			}
			if border.Color != "" {
				borderStyle.Color = "#" + strings.ToUpper(border.Color)
			}
			if border.Style != 0 {
				borderStyle.Style = intToBorderStyleName(border.Style)
			}
			borders = append(borders, borderStyle)
		}
		if len(borders) > 0 {
			result.Border = borders
		}
	}

	// Font
	if style.Font != nil {
		font := &FontStyle{}
		if style.Font.Bold {
			font.Bold = &style.Font.Bold
		}
		if style.Font.Italic {
			font.Italic = &style.Font.Italic
		}
		if style.Font.Underline != "" {
			underline := FontUnderline(style.Font.Underline)
			font.Underline = &underline
		}
		if style.Font.Size > 0 {
			size := int(style.Font.Size)
			font.Size = &size
		}
		if style.Font.Strike {
			font.Strike = &style.Font.Strike
		}
		if style.Font.Color != "" {
			color := "#" + strings.ToUpper(style.Font.Color)
			font.Color = &color
		}
		if style.Font.VertAlign != "" {
			vertAlign := FontVertAlign(style.Font.VertAlign)
			font.VertAlign = &vertAlign
		}
		if font.Bold != nil || font.Italic != nil || font.Underline != nil || font.Size != nil || font.Strike != nil || font.Color != nil || font.VertAlign != nil {
			result.Font = font
		}
	}

	// Fill
	if style.Fill.Type != "" || style.Fill.Pattern != 0 || len(style.Fill.Color) > 0 {
		fill := &FillStyle{}
		if style.Fill.Type != "" {
			fill.Type = FillType(style.Fill.Type)
		}
		if style.Fill.Pattern != 0 {
			fill.Pattern = intToFillPatternName(style.Fill.Pattern)
		}
		if len(style.Fill.Color) > 0 {
			var colors []string
			for _, color := range style.Fill.Color {
				if color != "" {
					colors = append(colors, "#"+strings.ToUpper(color))
				}
			}
			if len(colors) > 0 {
				fill.Color = colors
			}
		}
		if style.Fill.Shading != 0 {
			shading := intToFillShadingName(style.Fill.Shading)
			fill.Shading = &shading
		}
		if fill.Type != "" || fill.Pattern != FillPatternNone || len(fill.Color) > 0 || fill.Shading != nil {
			result.Fill = fill
		}
	}

	// NumFmt
	if style.CustomNumFmt != nil && *style.CustomNumFmt != "" {
		result.NumFmt = style.CustomNumFmt
	}

	// DecimalPlaces
	if style.DecimalPlaces != nil && *style.DecimalPlaces != 0 {
		result.DecimalPlaces = style.DecimalPlaces
	}

	return result
}

func intToBorderStyleName(style int) BorderStyle {
	styles := map[int]BorderStyle{
		0:  BorderStyleNone,
		1:  BorderStyleContinuous,
		2:  BorderStyleContinuous,
		3:  BorderStyleDash,
		4:  BorderStyleDot,
		5:  BorderStyleContinuous,
		6:  BorderStyleDouble,
		7:  BorderStyleContinuous,
		8:  BorderStyleDashDot,
		9:  BorderStyleDashDotDot,
		10: BorderStyleSlantDashDot,
		11: BorderStyleContinuous,
		12: BorderStyleMediumDashDot,
		13: BorderStyleMediumDashDotDot,
	}
	if name, exists := styles[style]; exists {
		return name
	}
	return BorderStyleContinuous
}

func intToFillPatternName(pattern int) FillPattern {
	patterns := map[int]FillPattern{
		0:  FillPatternNone,
		1:  FillPatternSolid,
		2:  FillPatternMediumGray,
		3:  FillPatternDarkGray,
		4:  FillPatternLightGray,
		5:  FillPatternDarkHorizontal,
		6:  FillPatternDarkVertical,
		7:  FillPatternDarkDown,
		8:  FillPatternDarkUp,
		9:  FillPatternDarkGrid,
		10: FillPatternDarkTrellis,
		11: FillPatternLightHorizontal,
		12: FillPatternLightVertical,
		13: FillPatternLightDown,
		14: FillPatternLightUp,
		15: FillPatternLightGrid,
		16: FillPatternLightTrellis,
		17: FillPatternGray125,
		18: FillPatternGray0625,
	}
	if name, exists := patterns[pattern]; exists {
		return name
	}
	return FillPatternNone
}

func intToFillShadingName(shading int) FillShading {
	shadings := map[int]FillShading{
		0: FillShadingHorizontal,
		1: FillShadingVertical,
		2: FillShadingDiagonalDown,
		3: FillShadingDiagonalUp,
		4: FillShadingFromCenter,
		5: FillShadingFromCorner,
	}
	if name, exists := shadings[shading]; exists {
		return name
	}
	return FillShadingHorizontal
}

func borderStyleNameToInt(style BorderStyle) int {
	styles := map[BorderStyle]int{
		BorderStyleNone:             0,
		BorderStyleContinuous:       1,
		BorderStyleDash:             3,
		BorderStyleDot:              4,
		BorderStyleDouble:           6,
		BorderStyleDashDot:          8,
		BorderStyleDashDotDot:       9,
		BorderStyleSlantDashDot:     10,
		BorderStyleMediumDashDot:    12,
		BorderStyleMediumDashDotDot: 13,
	}
	if value, exists := styles[style]; exists {
		return value
	}
	return 1
}

func fillPatternNameToInt(pattern FillPattern) int {
	patterns := map[FillPattern]int{
		FillPatternNone:            0,
		FillPatternSolid:           1,
		FillPatternMediumGray:      2,
		FillPatternDarkGray:        3,
		FillPatternLightGray:       4,
		FillPatternDarkHorizontal:  5,
		FillPatternDarkVertical:    6,
		FillPatternDarkDown:        7,
		FillPatternDarkUp:          8,
		FillPatternDarkGrid:        9,
		FillPatternDarkTrellis:     10,
		FillPatternLightHorizontal: 11,
		FillPatternLightVertical:   12,
		FillPatternLightDown:       13,
		FillPatternLightUp:         14,
		FillPatternLightGrid:       15,
		FillPatternLightTrellis:    16,
		FillPatternGray125:         17,
		FillPatternGray0625:        18,
	}
	if value, exists := patterns[pattern]; exists {
		return value
	}
	return 0
}

func fillShadingNameToInt(shading FillShading) int {
	shadings := map[FillShading]int{
		FillShadingHorizontal:   0,
		FillShadingVertical:     1,
		FillShadingDiagonalDown: 2,
		FillShadingDiagonalUp:   3,
		FillShadingFromCenter:   4,
		FillShadingFromCorner:   5,
	}
	if value, exists := shadings[shading]; exists {
		return value
	}
	return 0
}

// updateDimention updates the dimension of the worksheet after a cell is updated.
func (w *ExcelizeWorksheet) updateDimension(updatedCell string) error {
	dimension, err := w.file.GetSheetDimension(w.sheetName)
	if err != nil {
		return err
	}
	startCol, startRow, endCol, endRow, err := ParseRange(dimension)
	if err != nil {
		return err
	}
	updatedCol, updatedRow, err := excelize.CellNameToCoordinates(updatedCell)
	if err != nil {
		return err
	}
	if startCol > updatedCol {
		startCol = updatedCol
	}
	if endCol < updatedCol {
		endCol = updatedCol
	}
	if startRow > updatedRow {
		startRow = updatedRow
	}
	if endRow < updatedRow {
		endRow = updatedRow
	}
	startRange, err := excelize.CoordinatesToCellName(startCol, startRow)
	if err != nil {
		return err
	}
	endRange, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		return err
	}
	updatedDimension := fmt.Sprintf("%s:%s", startRange, endRange)
	return w.file.SetSheetDimension(w.sheetName, updatedDimension)
}

// Style creation helper functions for Excel formatting
// Excelフォーマット用のスタイル作成ヘルパー関数群だよ～ ✨(ﾉ◕ヮ◕)ﾉ*:･ﾟ✧

// createFontStyle creates font styling from format options
// フォーマットオプションからフォントスタイルを作成するのです！ (◕‿◕)♡
func createFontStyle(style map[string]interface{}) *excelize.Font {
	font := &excelize.Font{}
	
	if color, ok := style["fontColor"].(string); ok && color != "" {
		font.Color = strings.TrimPrefix(color, "#")
	}
	if name, ok := style["fontName"].(string); ok && name != "" {
		font.Family = name
	}
	if size, ok := style["fontSize"].(float64); ok && size > 0 {
		font.Size = size
	}
	if bold, ok := style["bold"].(bool); ok {
		font.Bold = bold
	}
	if italic, ok := style["italic"].(bool); ok {
		font.Italic = italic
	}
	if underline, ok := style["underline"].(bool); ok && underline {
		font.Underline = "single"
	}
	if strike, ok := style["strikethrough"].(bool); ok {
		font.Strike = strike
	}
	
	return font
}

// createFillStyle creates background fill styling
// 背景塗りつぶしスタイルを作成するのです！ 🎨(´∀｀)♡
func createFillStyle(style map[string]interface{}) excelize.Fill {
	fill := excelize.Fill{}
	
	if bgColor, ok := style["backgroundColor"].(string); ok && bgColor != "" {
		fill.Type = "pattern"
		fill.Pattern = 1 // solid fill
		fill.Color = []string{strings.TrimPrefix(bgColor, "#")}
	}
	
	return fill
}

// createAlignmentStyle creates text alignment styling  
// テキスト配置スタイルを作成するのです！ (◕‿◕)
func createAlignmentStyle(style map[string]interface{}) *excelize.Alignment {
	alignment := &excelize.Alignment{}
	
	if hAlign, ok := style["horizontalAlign"].(string); ok && hAlign != "" {
		alignment.Horizontal = hAlign
	}
	if vAlign, ok := style["verticalAlign"].(string); ok && vAlign != "" {
		alignment.Vertical = vAlign
	}
	if wrap, ok := style["wrapText"].(bool); ok {
		alignment.WrapText = wrap
	}
	
	return alignment
}

// createBorderStyle creates border styling
// ボーダースタイルを作成するのです！ ✨(｡◕‿◕｡)
func createBorderStyle(style map[string]interface{}) []excelize.Border {
	borders := []excelize.Border{}
	
	borderStyle, hasStyle := style["borderStyle"].(string)
	borderColor, hasColor := style["borderColor"].(string)
	
	if hasStyle && borderStyle != "" && borderStyle != "none" {
		color := "000000" // default black
		if hasColor && borderColor != "" {
			color = strings.TrimPrefix(borderColor, "#")
		}
		
		// Apply border to all sides - top, bottom, left, right
		// 上下左右全ての辺にボーダーを適用するのです！ ٩(◕‿◕)۶
		sides := []string{"top", "bottom", "left", "right"}
		for _, side := range sides {
			borders = append(borders, excelize.Border{
				Type:  side,
				Color: color,
				Style: getBorderStyleIndex(borderStyle),
			})
		}
	}
	
	return borders
}

// getBorderStyleIndex converts border style string to Excelize index
// ボーダースタイル文字列をExcelizeインデックスに変換するのです！ (´∀｀)
func getBorderStyleIndex(borderStyle string) int {
	switch borderStyle {
	case "thin":
		return 1
	case "medium":
		return 2
	case "thick":
		return 5
	case "double":
		return 6
	default:
		return 1 // default to thin
	}
}

// getNumberFormat extracts number format from style options and maps to Excel format codes
// スタイルオプションから数値フォーマットを抽出してExcelフォーマットコードにマッピングするのです！ ✨(◕‿◕)♡
func getNumberFormat(style map[string]interface{}) int {
	if numFormat, ok := style["numberFormat"].(string); ok && numFormat != "" {
		// Map common number format strings to Excel built-in format codes
		// 一般的な数値フォーマット文字列をExcelの組み込みフォーマットコードにマッピングするのです！ (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
		formatMap := map[string]int{
			// Basic number formats
			// 基本的な数値フォーマット ٩(◕‿◕)۶
			"General":    0,
			"0":          1,
			"0.0":        2,
			"0.00":       2,
			"#,##0":      3,
			"#,##0.0":    4,
			"#,##0.00":   4,
			
			// Currency formats
			// 通貨フォーマット 💰(´∀｀)♡
			"$#,##0":       5,
			"$#,##0.0":     6,
			"$#,##0.00":    7,
			"$#,##0.00_);[Red]($#,##0.00)": 8,
			"0%":           9,
			"0.0%":         10,
			"0.00%":        10,
			
			// Date formats  
			// 日付フォーマット 📅✨
			"m/d/yyyy":      14,
			"d-mmm-yy":      15,
			"d-mmm":         16,
			"mmm-yy":        17,
			"h:mm AM/PM":    18,
			"h:mm:ss AM/PM": 19,
			"h:mm":          20,
			"h:mm:ss":       21,
			"m/d/yyyy h:mm": 22,
			"mm/dd/yyyy":    14,
			"dd/mm/yyyy":    14,
			"yyyy-mm-dd":    14,
			"yyyy/mm/dd":    14,
			
			// Time formats
			// 時刻フォーマット ⏰(◕‿◕)
			"mm:ss":         45,
			"[h]:mm:ss":     46,
			"mm:ss.0":       47,
			"@":             49, // Text format
			
			// Extended currency formats
			// 拡張通貨フォーマット 💸
			"¥#,##0":       164, // CN¥ 
			"€#,##0.00":    164, // Euro
			"£#,##0.00":    164, // Pound
			"₹#,##0.00":    164, // Rupee
			
			// Percentage with decimals
			// 小数点付きパーセンテージ 📊
			"#,##0%":       9,
			"#,##0.0%":     10,
			"#,##0.00%":    10,
			
			// Scientific notation  
			// 科学記数法 🔬(´∀｀)
			"0.00E+00":     11,
			"##0.0E+0":     48,
			
			// Fractions
			// 分数 ➗✨
			"# ?/?":        12,
			"# ??/??":      13,
			"# ???/???":    13,
			
			// Accounting formats
			// 会計フォーマット 💼
			"_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)": 37,
			"_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)": 38,
			"_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)": 39,
			"_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)": 40,
		}
		
		// Check if we have a direct mapping
		// 直接マッピングがあるかチェックするのです！ (◕‿◕)
		if code, exists := formatMap[numFormat]; exists {
			return code
		}
		
		// Try to match partial patterns for common formats
		// 一般的なフォーマットの部分パターンをマッチしてみるのです！ ✨
		switch {
		case strings.Contains(numFormat, "$") && strings.Contains(numFormat, "#,##0.00"):
			return 7 // Currency with 2 decimals
		case strings.Contains(numFormat, "$") && strings.Contains(numFormat, "#,##0"):
			return 5 // Currency with no decimals  
		case strings.Contains(numFormat, "%"):
			if strings.Contains(numFormat, ".00") {
				return 10 // Percentage with 2 decimals
			} else if strings.Contains(numFormat, ".0") {
				return 10 // Percentage with 1 decimal
			}
			return 9 // Simple percentage
		case strings.Contains(numFormat, "#,##0.00"):
			return 4 // Number with thousands separator and 2 decimals
		case strings.Contains(numFormat, "#,##0"):
			return 3 // Number with thousands separator
		case strings.Contains(numFormat, "0.00"):
			return 2 // Number with 2 decimals
		case strings.Contains(numFormat, "yyyy") || strings.Contains(numFormat, "mm") || strings.Contains(numFormat, "dd"):
			return 14 // Date format
		case strings.Contains(numFormat, "h:mm"):
			return 20 // Time format
		case numFormat == "@":
			return 49 // Text format
		}
	}
	
	return 0 // Default to General format / 一般形式をデフォルトにするのです！ (´∀｀)♡
}

// parseRangeForExcelize parses a range string and returns start and end cells for Excelize
// Excelize用に範囲文字列を解析して開始セルと終了セルを返すのです！ ✨(ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
func parseRangeForExcelize(rangeStr string) (string, string, error) {
	// Handle single cell references (e.g., "A1" -> "A1", "A1")
	// 単一セル参照を処理するのです (e.g., "A1" -> "A1", "A1") (◕‿◕)♡
	if !strings.Contains(rangeStr, ":") {
		return rangeStr, rangeStr, nil
	}
	
	// Handle range references (e.g., "A1:C3" -> "A1", "C3")
	// 範囲参照を処理するのです (e.g., "A1:C3" -> "A1", "C3") ✨
	parts := strings.Split(rangeStr, ":")
	if len(parts) != 2 {
		return "", "", fmt.Errorf("invalid range format: %s", rangeStr)
	}
	
	return strings.TrimSpace(parts[0]), strings.TrimSpace(parts[1]), nil
}
