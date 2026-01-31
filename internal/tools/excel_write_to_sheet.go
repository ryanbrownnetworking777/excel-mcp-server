package tools

import (
	"context"
	"fmt"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	"github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
	"github.com/xuri/excelize/v2"
)

// ExcelWriteToSheetArguments defines the structure for writing Excel data
// Excel書き込み引数の構造体だよ〜 (◕‿◕)✨
type ExcelWriteToSheetArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	NewSheet         bool   `zog:"newSheet"`
	Range            string `zog:"range"`
	// Values handled manually due to mixed type support
	// 混合型サポートのため手動処理するのです！(｡◕‿‿◕｡)
}

// Schema validation for Excel write arguments
// Excelの書き込み引数のスキーマバリデーション〜 ٩(◕‿◕)۶
var excelWriteToSheetArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"newSheet":         z.Bool().Required().Default(false),
	"range":            z.String().Required(),
	// values omitted - handled manually to support mixed types as promised in MCP schema
	// valuesは省略 - MCPスキーマで約束した混合型をサポートするため手動処理なのです！ ╰( ͡° ͜ʖ ͡° )つ──☆*:・ﾟ
})

func AddExcelWriteToSheetTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_write_to_sheet",
		mcp.WithDescription("Write values to the Excel sheet"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithBoolean("newSheet",
			mcp.Required(),
			mcp.Description("Create a new sheet if true, otherwise write to the existing sheet"),
		),
		mcp.WithString("range",
			mcp.Required(),
			mcp.Description("Range of cells in the Excel sheet (e.g., \"A1:C10\")"),
		),
		mcp.WithArray("values",
			mcp.Required(),
			mcp.Description("Values to write to the Excel sheet. If the value is a formula, it should start with \"=\""),
			mcp.Items(map[string]any{
				"type": "array",
				"items": map[string]any{
					"anyOf": []any{
						map[string]any{
							"type": "string",
						},
						map[string]any{
							"type": "number",
						},
						map[string]any{
							"type": "boolean",
						},
						map[string]any{
							"type": "null",
						},
					},
				},
			}),
		),
	), handleWriteToSheet)
}

func handleWriteToSheet(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelWriteToSheetArguments{}
	issues := excelWriteToSheetArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}

	// Handle values manually to support mixed types (string, number, boolean, null)
	// as promised in the MCP tool schema - this fixes Claude Desktop compatibility!
	// 混合型（文字列、数値、ブール値、null）をサポートするため手動処理です〜
	// MCPツールスキーマで約束したとおりに！Claude Desktopとの互換性を修正するのです！ (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
	valuesArray, ok := request.GetArguments()["values"].([]any)
	if !ok {
		return imcp.NewToolResultInvalidArgumentError("values must be a 2D array"), nil
	}
	
	// Convert to proper 2D array format that matches our MCP schema promise
	// MCPスキーマの約束に合う適切な2D配列形式に変換だー！ ヽ(°〇°)ﾉ
	values := make([][]any, len(valuesArray))
	for i, rowAny := range valuesArray {
		row, ok := rowAny.([]any)
		if !ok {
			return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("values[%d] must be an array", i)), nil
		}
		values[i] = row
	}

	return writeSheet(args.FileAbsolutePath, args.SheetName, args.NewSheet, args.Range, values)
}

func writeSheet(fileAbsolutePath string, sheetName string, newSheet bool, rangeStr string, values [][]any) (*mcp.CallToolResult, error) {
	workbook, closeFn, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer closeFn()

	startCol, startRow, endCol, endRow, err := excel.ParseRange(rangeStr)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}

	// データの整合性チェック
	rangeRowSize := endRow - startRow + 1
	if len(values) != rangeRowSize {
		return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("number of rows in data (%d) does not match range size (%d)", len(values), rangeRowSize)), nil
	}

	if newSheet {
		if err := workbook.CreateNewSheet(sheetName); err != nil {
			return nil, err
		}
	}

	// シートの取得
	worksheet, err := workbook.FindSheet(sheetName)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	defer worksheet.Release()

	// データの書き込み
	wroteFormula := false
	var formulaCells []string // Track formula cells for validation
	// 数式セルの検証用トラッキング配列だよ～ (◕‿◕)♡
	
	for i, row := range values {
		rangeColumnSize := endCol - startCol + 1
		if len(row) != rangeColumnSize {
			return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("number of columns in row %d (%d) does not match range size (%d)", i, len(row), rangeColumnSize)), nil
		}
		for j, cellValue := range row {
			cell, err := excelize.CoordinatesToCellName(startCol+j, startRow+i)
			if err != nil {
				return nil, err
			}
			if cellStr, ok := cellValue.(string); ok && isFormula(cellStr) {
				// Formula detected! Time for some spreadsheet calculations! 📊
				// 数式発見！スプレッドシートの計算タイムです！ 📊(*＾▽＾*)
				err = worksheet.SetFormula(cell, cellStr)
				wroteFormula = true
				formulaCells = append(formulaCells, cell) // Track for validation
			} else {
				// Regular value - just write it directly! Simple and clean~
				// 普通の値 - そのまま書き込むだけ！シンプルでキレイ～ (´∀｀)♡
				err = worksheet.SetValue(cell, cellValue)
			}
			if err != nil {
				return nil, err
			}
		}
	}

	if err := workbook.Save(); err != nil {
		return nil, err
	}

	// Get formula results and validate for errors after saving! 🔍
	// 保存後に数式結果を取得してエラーをチェックするのです！ 🔍(｡◕‿◕｡)
	var formulaResults []FormulaResult
	var formulaErrors []FormulaError
	if len(formulaCells) > 0 {
		formulaResults = getFormulaResults(worksheet, formulaCells)
		formulaErrors = validateFormulas(worksheet, formulaCells)
	}

	// HTMLテーブルの生成
	var table *string
	if wroteFormula {
		table, err = CreateHTMLTableOfFormula(worksheet, startCol, startRow, endCol, endRow)
	} else {
		table, err = CreateHTMLTableOfValues(worksheet, startCol, startRow, endCol, endRow)
	}
	if err != nil {
		return nil, err
	}
	html := "<h2>Written Sheet</h2>\n"
	html += *table + "\n"
	html += "<h2>Metadata</h2>\n"
	html += "<ul>\n"
	html += fmt.Sprintf("<li>backend: %s</li>\n", workbook.GetBackendName())
	html += fmt.Sprintf("<li>sheet name: %s</li>\n", sheetName)
	html += fmt.Sprintf("<li>write range: %s</li>\n", rangeStr)
	if len(formulaCells) > 0 {
		html += fmt.Sprintf("<li>formulas written: %d</li>\n", len(formulaCells))
	}
	html += "</ul>\n"
	
	// Add formula results for Claude Desktop feedback! Critical for understanding formula output! 🔍
	// Claude Desktopのフィードバック用数式結果を追加！数式の出力を理解するのに重要です！ 🔍✨
	if len(formulaResults) > 0 {
		html += "<h2>🔢 Formula Results</h2>\n"
		html += "<div style='background-color: #f8f9fa; border: 1px solid #dee2e6; padding: 15px; margin: 10px 0;'>\n"
		html += "<p><strong>Calculated values for all formulas:</strong></p>\n"
		html += "<table style='border-collapse: collapse; width: 100%; margin: 10px 0;'>\n"
		html += "<tr style='background-color: #e9ecef;'>\n"
		html += "<th style='border: 1px solid #adb5bd; padding: 8px; text-align: left;'>Cell</th>\n"
		html += "<th style='border: 1px solid #adb5bd; padding: 8px; text-align: left;'>Formula</th>\n"
		html += "<th style='border: 1px solid #adb5bd; padding: 8px; text-align: left;'>Result</th>\n"
		html += "<th style='border: 1px solid #adb5bd; padding: 8px; text-align: center;'>Status</th>\n"
		html += "</tr>\n"
		
		for _, result := range formulaResults {
			statusColor := "green"
			statusIcon := "✅"
			if result.IsError {
				statusColor = "red"
				statusIcon = "❌"
			}
			
			html += "<tr>\n"
			html += fmt.Sprintf("<td style='border: 1px solid #adb5bd; padding: 8px;'><strong>%s</strong></td>\n", result.Cell)
			html += fmt.Sprintf("<td style='border: 1px solid #adb5bd; padding: 8px;'><code>%s</code></td>\n", result.Formula)
			html += fmt.Sprintf("<td style='border: 1px solid #adb5bd; padding: 8px; color: %s;'><strong>%s</strong></td>\n", statusColor, result.Value)
			html += fmt.Sprintf("<td style='border: 1px solid #adb5bd; padding: 8px; text-align: center;'>%s</td>\n", statusIcon)
			html += "</tr>\n"
		}
		html += "</table>\n"
		html += "</div>\n"
	}
	
	// Add detailed error information if any formulas failed
	// エラーがある場合は詳細なエラー情報を追加するのです！ (╯°□°）╯
	if len(formulaErrors) > 0 {
		html += "<h2>⚠️ Formula Error Details</h2>\n"
		html += "<div style='background-color: #fff3cd; border: 1px solid #ffeaa7; padding: 15px; margin: 10px 0;'>\n"
		html += "<p><strong>Detailed error explanations:</strong></p>\n"
		html += "<ul>\n"
		for _, ferr := range formulaErrors {
			html += fmt.Sprintf("<li><strong>Cell %s</strong>: <code>%s</code><br>", ferr.Cell, ferr.Formula)
			html += fmt.Sprintf("   <span style='color: red;'>%s</span></li>\n", ferr.Error)
		}
		html += "</ul>\n"
		html += "<p><em>Please review and correct these formulas before proceeding.</em></p>\n"
		html += "</div>\n"
	}
	
	html += "<h2>Notice</h2>\n"
	if len(formulaErrors) > 0 {
		html += fmt.Sprintf("<p>⚠️ Data written successfully, but %d formula(s) contain errors. Please review the errors above.</p>\n", len(formulaErrors))
	} else {
		html += "<p>✅ Values and formulas written successfully.</p>\n"
	}

	return mcp.NewToolResultText(html), nil
}

func isFormula(value string) bool {
	return len(value) > 0 && value[0] == '='
}

// FormulaError represents a formula validation error
// 数式検証エラーを表す構造体だよ～ (╯°□°）╯
type FormulaError struct {
	Cell    string `json:"cell"`
	Formula string `json:"formula"`
	Error   string `json:"error"`
	Value   string `json:"value"`
}

// FormulaResult represents a formula and its calculated value
// 数式とその計算結果を表す構造体です！ ٩(◕‿◕)۶
type FormulaResult struct {
	Cell    string `json:"cell"`
	Formula string `json:"formula"`
	Value   string `json:"value"`
	IsError bool   `json:"isError"`
}

// getFormulaResults retrieves all formula results for Claude Desktop feedback
// Claude Desktopのフィードバック用に全ての数式結果を取得するのです！ ✨(ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
func getFormulaResults(worksheet excel.Worksheet, formulaCells []string) []FormulaResult {
	var results []FormulaResult
	
	for _, cell := range formulaCells {
		// Get the formula
		formula, err := worksheet.GetFormula(cell)
		if err != nil {
			results = append(results, FormulaResult{
				Cell:    cell,
				Formula: "unknown",
				Value:   "ERROR",
				IsError: true,
			})
			continue
		}
		
		// Get the calculated value
		value, err := worksheet.GetValue(cell)
		if err != nil {
			results = append(results, FormulaResult{
				Cell:    cell,
				Formula: formula,
				Value:   "ERROR",
				IsError: true,
			})
			continue
		}
		
		// Check if it's an Excel error
		isError := isExcelError(value)
		
		results = append(results, FormulaResult{
			Cell:    cell,
			Formula: formula,
			Value:   value,
			IsError: isError,
		})
	}
	
	return results
}

// validateFormulas checks for Excel errors in formula cells
// 数式セルのExcelエラーをチェックする関数です！ ٩(◕‿◕)۶
func validateFormulas(worksheet excel.Worksheet, formulaCells []string) []FormulaError {
	var errors []FormulaError
	
	for _, cell := range formulaCells {
		// Get the formula
		formula, err := worksheet.GetFormula(cell)
		if err != nil {
			errors = append(errors, FormulaError{
				Cell:    cell,
				Formula: "unknown",
				Error:   fmt.Sprintf("Failed to get formula: %v", err),
				Value:   "ERROR",
			})
			continue
		}
		
		// Get the calculated value to check for Excel errors
		value, err := worksheet.GetValue(cell)
		if err != nil {
			errors = append(errors, FormulaError{
				Cell:    cell,
				Formula: formula,
				Error:   fmt.Sprintf("Failed to calculate: %v", err),
				Value:   "ERROR",
			})
			continue
		}
		
		// Check for common Excel error values
		if isExcelError(value) {
			errors = append(errors, FormulaError{
				Cell:    cell,
				Formula: formula,
				Error:   getExcelErrorDescription(value),
				Value:   value,
			})
		}
	}
	
	return errors
}

// isExcelError checks if a value represents an Excel error
// 値がExcelエラーかどうかチェックするのです！ (°o°)
func isExcelError(value string) bool {
	excelErrors := []string{
		"#DIV/0!",   // Division by zero
		"#N/A",      // Value not available  
		"#NAME?",    // Name error
		"#NULL!",    // Null error
		"#NUM!",     // Number error
		"#REF!",     // Reference error
		"#VALUE!",   // Value error
		"#GETTING_DATA", // Getting data (newer Excel)
	}
	
	for _, errorCode := range excelErrors {
		if value == errorCode {
			return true
		}
	}
	return false
}

// getExcelErrorDescription provides human-readable error descriptions
// 人間が読みやすいエラー説明を提供するのです！ (´∀｀)
func getExcelErrorDescription(errorCode string) string {
	descriptions := map[string]string{
		"#DIV/0!":        "Division by zero - check for empty cells or zero values in denominators",
		"#N/A":           "Value not available - function cannot find referenced data",
		"#NAME?":         "Name not recognized - check function names and cell references",
		"#NULL!":         "Null error - invalid intersection of ranges",
		"#NUM!":          "Number error - invalid numeric values or arguments",
		"#REF!":          "Reference error - invalid cell reference",
		"#VALUE!":        "Value error - wrong data type for function",
		"#GETTING_DATA":  "Still calculating - data may not be fully loaded",
	}
	
	if desc, exists := descriptions[errorCode]; exists {
		return desc
	}
	return fmt.Sprintf("Unknown Excel error: %s", errorCode)
}
