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
var excelWriteToSheetArgumentsSchema = z.Struct(z.Schema{
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
	valuesArg, ok := request.Params.Arguments["values"]
	if !ok {
		return imcp.NewToolResultInvalidArgumentError("missing required parameter: values"), nil
	}
	
	valuesArray, ok := valuesArg.([]any)
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

	startCol, startRow, endCol, endRow, err := excel.ParseRange(rangeStr)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}

	// Data integrity check - making sure dimensions match perfectly!
	// データの整合性チェック - 寸法がピッタリ合うか確認するのです！ (｡･ω･｡)ﾉ♡
	rangeRowSize := endRow - startRow + 1
	if len(values) != rangeRowSize {
		return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("number of rows in data (%d) does not match range size (%d)", len(values), rangeRowSize)), nil
	}

	// Time to write the data! Let's make some Excel magic happen ✨
	// データの書き込みタイム！Excelの魔法を発動させるのです！ ✨(ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
	wroteFormula := false
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
	html += fmt.Sprintf("<li>read range: %s</li>\n", rangeStr)
	html += "</ul>\n"
	html += "<h2>Notice</h2>\n"
	html += "<p>Values wrote successfully.</p>\n"

	return mcp.NewToolResultText(html), nil
}

func isFormula(value string) bool {
	return len(value) > 0 && value[0] == '='
}
