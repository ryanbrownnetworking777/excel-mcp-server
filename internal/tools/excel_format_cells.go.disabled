package tools

import (
	"context"
	"fmt"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	"github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

// ExcelFormatCellsArguments defines the structure for Excel cell formatting
// Excelセルフォーマットの構造体だよ～ ✨(ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
type ExcelFormatCellsArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Range            string `zog:"range"`
	// Formatting options handled manually for flexibility
}

// CellFormat represents comprehensive Excel cell formatting options
// 包括的なExcelセルフォーマットオプションを表すのです！ (◕‿◕)♡
type CellFormat struct {
	// Colors
	BackgroundColor string `json:"backgroundColor,omitempty"` // Hex color like "#FF0000"
	FontColor       string `json:"fontColor,omitempty"`       // Hex color like "#000000"
	
	// Font styling
	FontName   string `json:"fontName,omitempty"`   // e.g., "Arial", "Calibri"
	FontSize   int    `json:"fontSize,omitempty"`   // e.g., 12, 14
	Bold       bool   `json:"bold,omitempty"`
	Italic     bool   `json:"italic,omitempty"`
	Underline  bool   `json:"underline,omitempty"`
	
	// Number formatting
	NumberFormat string `json:"numberFormat,omitempty"` // e.g., "0.00", "#,##0", "mm/dd/yyyy"
	
	// Alignment
	HorizontalAlign string `json:"horizontalAlign,omitempty"` // "left", "center", "right"
	VerticalAlign   string `json:"verticalAlign,omitempty"`   // "top", "middle", "bottom"
	WrapText        bool   `json:"wrapText,omitempty"`
	
	// Borders
	BorderStyle string `json:"borderStyle,omitempty"` // "thin", "thick", "medium", "double"
	BorderColor string `json:"borderColor,omitempty"` // Hex color
	
	// Special formatting
	Strikethrough bool `json:"strikethrough,omitempty"`
}

var excelFormatCellsArgumentsSchema = z.Struct(z.Schema{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"range":            z.String().Required(),
	// format options handled manually for flexibility
})

func AddExcelFormatCellsTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_format_cells",
		mcp.WithDescription("Apply comprehensive formatting to Excel cells including colors, fonts, numbers, alignment, and borders"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("range",
			mcp.Required(),
			mcp.Description("Range of cells to format (e.g., \"A1:C10\" or \"A1\")"),
		),
		mcp.WithObject("format",
			mcp.Required(),
			mcp.Description("Formatting options to apply"),
			mcp.Properties(map[string]any{
				"backgroundColor": map[string]any{
					"type":        "string",
					"description": "Background color in hex format (e.g., \"#FF0000\" for red)",
				},
				"fontColor": map[string]any{
					"type":        "string",
					"description": "Font color in hex format (e.g., \"#000000\" for black)",
				},
				"fontName": map[string]any{
					"type":        "string",
					"description": "Font family name (e.g., \"Arial\", \"Calibri\", \"Times New Roman\")",
				},
				"fontSize": map[string]any{
					"type":        "integer",
					"description": "Font size in points (e.g., 12, 14, 16)",
				},
				"bold": map[string]any{
					"type":        "boolean",
					"description": "Apply bold formatting",
				},
				"italic": map[string]any{
					"type":        "boolean",
					"description": "Apply italic formatting",
				},
				"underline": map[string]any{
					"type":        "boolean",
					"description": "Apply underline formatting",
				},
				"numberFormat": map[string]any{
					"type":        "string",
					"description": "Number format string (e.g., \"0.00\", \"#,##0\", \"mm/dd/yyyy\", \"$#,##0.00\")",
				},
				"horizontalAlign": map[string]any{
					"type":        "string",
					"enum":        []string{"left", "center", "right"},
					"description": "Horizontal text alignment",
				},
				"verticalAlign": map[string]any{
					"type":        "string",
					"enum":        []string{"top", "middle", "bottom"},
					"description": "Vertical text alignment",
				},
				"wrapText": map[string]any{
					"type":        "boolean",
					"description": "Enable text wrapping",
				},
				"borderStyle": map[string]any{
					"type":        "string",
					"enum":        []string{"thin", "thick", "medium", "double", "none"},
					"description": "Border style for all borders",
				},
				"borderColor": map[string]any{
					"type":        "string",
					"description": "Border color in hex format",
				},
				"strikethrough": map[string]any{
					"type":        "boolean",
					"description": "Apply strikethrough formatting",
				},
			}),
		),
	), handleFormatCells)
}

func handleFormatCells(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelFormatCellsArguments{}
	issues := excelFormatCellsArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}

	// Handle format options manually for full flexibility
	// 完全な柔軟性のためフォーマットオプションを手動処理するのです！ ٩(◕‿◕)۶
	formatArg, ok := request.Params.Arguments["format"]
	if !ok {
		return imcp.NewToolResultInvalidArgumentError("missing required parameter: format"), nil
	}
	
	formatMap, ok := formatArg.(map[string]any)
	if !ok {
		return imcp.NewToolResultInvalidArgumentError("format must be an object"), nil
	}
	
	// Parse format options
	format := parseCellFormat(formatMap)
	
	return formatCells(args.FileAbsolutePath, args.SheetName, args.Range, format)
}

// parseCellFormat converts the JSON format object to our CellFormat struct
// JSONフォーマットオブジェクトをCellFormat構造体に変換するのです！ (´∀｀)♡
func parseCellFormat(formatMap map[string]any) CellFormat {
	format := CellFormat{}
	
	if val, ok := formatMap["backgroundColor"].(string); ok {
		format.BackgroundColor = val
	}
	if val, ok := formatMap["fontColor"].(string); ok {
		format.FontColor = val
	}
	if val, ok := formatMap["fontName"].(string); ok {
		format.FontName = val
	}
	if val, ok := formatMap["fontSize"].(float64); ok {
		format.FontSize = int(val)
	}
	if val, ok := formatMap["bold"].(bool); ok {
		format.Bold = val
	}
	if val, ok := formatMap["italic"].(bool); ok {
		format.Italic = val
	}
	if val, ok := formatMap["underline"].(bool); ok {
		format.Underline = val
	}
	if val, ok := formatMap["numberFormat"].(string); ok {
		format.NumberFormat = val
	}
	if val, ok := formatMap["horizontalAlign"].(string); ok {
		format.HorizontalAlign = val
	}
	if val, ok := formatMap["verticalAlign"].(string); ok {
		format.VerticalAlign = val
	}
	if val, ok := formatMap["wrapText"].(bool); ok {
		format.WrapText = val
	}
	if val, ok := formatMap["borderStyle"].(string); ok {
		format.BorderStyle = val
	}
	if val, ok := formatMap["borderColor"].(string); ok {
		format.BorderColor = val
	}
	if val, ok := formatMap["strikethrough"].(bool); ok {
		format.Strikethrough = val
	}
	
	return format
}

func formatCells(fileAbsolutePath string, sheetName string, rangeStr string, format CellFormat) (*mcp.CallToolResult, error) {
	workbook, closeFn, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer closeFn()

	// Get the worksheet
	worksheet, err := workbook.FindSheet(sheetName)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	defer worksheet.Release()

	// Apply formatting based on backend type
	// バックエンドタイプに基づいてフォーマットを適用するのです！ ✨(ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
	err = applyFormatting(workbook, worksheet, rangeStr, format)
	if err != nil {
		return nil, fmt.Errorf("failed to apply formatting: %w", err)
	}

	if err := workbook.Save(); err != nil {
		return nil, err
	}

	// Generate response showing applied formatting
	// 適用されたフォーマットを表示するレスポンスを生成するのです！ (◕‿◕)♡
	html := "<h2>✨ Cell Formatting Applied</h2>\n"
	html += fmt.Sprintf("<p><strong>File:</strong> %s</p>\n", fileAbsolutePath)
	html += fmt.Sprintf("<p><strong>Sheet:</strong> %s</p>\n", sheetName)
	html += fmt.Sprintf("<p><strong>Range:</strong> %s</p>\n", rangeStr)
	html += fmt.Sprintf("<p><strong>Backend:</strong> %s</p>\n", workbook.GetBackendName())
	
	html += "<h3>🎨 Formatting Applied:</h3>\n"
	html += "<ul>\n"
	
	if format.BackgroundColor != "" {
		html += fmt.Sprintf("<li>Background Color: <span style='background-color: %s; padding: 2px 8px; color: white;'>%s</span></li>\n", format.BackgroundColor, format.BackgroundColor)
	}
	if format.FontColor != "" {
		html += fmt.Sprintf("<li>Font Color: <span style='color: %s;'>%s</span></li>\n", format.FontColor, format.FontColor)
	}
	if format.FontName != "" {
		html += fmt.Sprintf("<li>Font: %s</li>\n", format.FontName)
	}
	if format.FontSize > 0 {
		html += fmt.Sprintf("<li>Font Size: %dpt</li>\n", format.FontSize)
	}
	if format.Bold {
		html += "<li><strong>Bold</strong> formatting applied</li>\n"
	}
	if format.Italic {
		html += "<li><em>Italic</em> formatting applied</li>\n"
	}
	if format.Underline {
		html += "<li><u>Underline</u> formatting applied</li>\n"
	}
	if format.NumberFormat != "" {
		html += fmt.Sprintf("<li>Number Format: <code>%s</code></li>\n", format.NumberFormat)
	}
	if format.HorizontalAlign != "" {
		html += fmt.Sprintf("<li>Horizontal Alignment: %s</li>\n", format.HorizontalAlign)
	}
	if format.VerticalAlign != "" {
		html += fmt.Sprintf("<li>Vertical Alignment: %s</li>\n", format.VerticalAlign)
	}
	if format.WrapText {
		html += "<li>Text Wrapping enabled</li>\n"
	}
	if format.BorderStyle != "" {
		html += fmt.Sprintf("<li>Border Style: %s", format.BorderStyle)
		if format.BorderColor != "" {
			html += fmt.Sprintf(" (Color: %s)", format.BorderColor)
		}
		html += "</li>\n"
	}
	if format.Strikethrough {
		html += "<li><s>Strikethrough</s> formatting applied</li>\n"
	}
	
	html += "</ul>\n"
	html += "<p>✅ Formatting has been successfully applied to the specified range!</p>\n"

	return mcp.NewToolResultText(html), nil
}

// applyFormatting applies the formatting based on the Excel backend type  
// Excelバックエンドタイプに基づいてフォーマットを適用するのです！ ✨(｡◕‿◕｡)
func applyFormatting(workbook excel.Excel, worksheet excel.Worksheet, rangeStr string, format CellFormat) error {
	// Convert our CellFormat to a map for the interface
	// インターフェース用にCellFormatをmapに変換するのです！ (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
	styleMap := map[string]interface{}{
		"backgroundColor":  format.BackgroundColor,
		"fontColor":       format.FontColor, 
		"fontName":        format.FontName,
		"fontSize":        float64(format.FontSize),
		"bold":            format.Bold,
		"italic":          format.Italic,
		"underline":       format.Underline,
		"numberFormat":    format.NumberFormat,
		"horizontalAlign": format.HorizontalAlign,
		"verticalAlign":   format.VerticalAlign,
		"wrapText":        format.WrapText,
		"borderStyle":     format.BorderStyle,
		"borderColor":     format.BorderColor,
		"strikethrough":   format.Strikethrough,
	}
	
	// Use the new FormatCells interface method
	// 新しいFormatCellsインターフェースメソッドを使用するのです！ ✨(◕‿◕)♡
	sheetName, err := worksheet.Name()
	if err != nil {
		return fmt.Errorf("failed to get sheet name: %w", err)
	}
	
	return workbook.FormatCells(sheetName, rangeStr, styleMap)
}

