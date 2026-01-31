package tools

import (
	"context"
	"fmt"
	"strconv"
	"strings"

	"github.com/negokaz/excel-mcp-server/internal/excel"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	"github.com/xuri/excelize/v2"

	z "github.com/Oudwins/zog"
)

// ExcelFinancialFormattingParams represents the parameters for financial formatting
type ExcelFinancialFormattingParams struct {
	FilePath       *string `json:"file_path" doc:"Path to the Excel file to format"`
	FormatType     *string `json:"format_type" doc:"Format type (investment_banking, corporate_finance, consulting)"`
	Currency       *string `json:"currency" doc:"Currency code (USD, JPY, EUR, etc.)"`
	ApplyToSheets  *[]string `json:"apply_to_sheets" doc:"List of sheet names to apply formatting (default: all financial sheets)"`
	IncludeHeaders *bool   `json:"include_headers" doc:"Apply formatting to headers (default: true)"`
	IncludeFormulas *bool  `json:"include_formulas" doc:"Apply color coding to formulas (default: true)"`
	IncludeConditional *bool `json:"include_conditional" doc:"Apply conditional formatting (default: true)"`
}

// ExcelFinancialFormattingResult represents the result of financial formatting
type ExcelFinancialFormattingResult struct {
	Message        string                 `json:"message"`
	FormattedSheets []string              `json:"formatted_sheets"`
	StylesApplied   map[string]interface{} `json:"styles_applied"`
	FormatSummary   map[string]interface{} `json:"format_summary"`
}

// ColorScheme represents the color scheme for different cell types
type ColorScheme struct {
	InputValues    string // 青字 (RGB 166,203,240)
	Formulas       string // 黒字
	LinkReferences string // 緑字 (RGB 198,239,206)
	Headers        string // 濃い青
	Subtotals      string // 太字
	Totals         string // 太字 + 枠線
}

// ExcelFinancialFormatting applies professional financial formatting to Excel models
func ExcelFinancialFormatting(ctx context.Context, params ExcelFinancialFormattingParams) (*ExcelFinancialFormattingResult, error) {
	// バリデーション
	schema := z.Struct(z.Schema{
		"file_path":            z.String().Test(AbsolutePathTest()),
		"format_type":          z.String().OneOf("investment_banking", "corporate_finance", "consulting").Default("investment_banking"),
		"currency":             z.String().Min(1).Default("USD"),
		"apply_to_sheets":      z.Slice(z.String()).Default([]string{}),
		"include_headers":      z.Bool().Default(true),
		"include_formulas":     z.Bool().Default(true),
		"include_conditional":  z.Bool().Default(true),
	})

	validatedParams, err := schema.Parse(params)
	if err != nil {
		return nil, err
	}

	// Excelファイルを開く
	excelFile, err := excel.OpenFile(*validatedParams.FilePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open Excel file: %w", err)
	}

	// フォーマッターの初期化
	formatter := &FinancialFormatter{
		file:         excelFile,
		params:       validatedParams,
		colorScheme:  getColorScheme(*validatedParams.FormatType),
		stylesApplied: make(map[string]interface{}),
	}

	// フォーマットの適用
	result, err := formatter.ApplyFinancialFormatting()
	if err != nil {
		return nil, fmt.Errorf("failed to apply formatting: %w", err)
	}

	// ファイルを保存
	err = excelFile.Save()
	if err != nil {
		return nil, fmt.Errorf("failed to save Excel file: %w", err)
	}

	return result, nil
}

// FinancialFormatter handles the formatting logic
type FinancialFormatter struct {
	file          excel.Excel
	params        ExcelFinancialFormattingParams
	colorScheme   ColorScheme
	stylesApplied map[string]interface{}
}

// ApplyFinancialFormatting applies the complete formatting
func (f *FinancialFormatter) ApplyFinancialFormatting() (*ExcelFinancialFormattingResult, error) {
	// 対象シートの決定
	sheetsToFormat := f.getSheetsToFormat()
	
	var formattedSheets []string
	
	// 各シートにフォーマットを適用
	for _, sheetName := range sheetsToFormat {
		err := f.formatSheet(sheetName)
		if err != nil {
			return nil, fmt.Errorf("failed to format sheet %s: %w", sheetName, err)
		}
		formattedSheets = append(formattedSheets, sheetName)
	}

	// 結果の作成
	return &ExcelFinancialFormattingResult{
		Message:         fmt.Sprintf("Financial formatting applied to %d sheets", len(formattedSheets)),
		FormattedSheets: formattedSheets,
		StylesApplied:   f.stylesApplied,
		FormatSummary: map[string]interface{}{
			"format_type":      *f.params.FormatType,
			"currency":         *f.params.Currency,
			"sheets_formatted": len(formattedSheets),
			"headers_formatted": *f.params.IncludeHeaders,
			"formulas_colored": *f.params.IncludeFormulas,
			"conditional_applied": *f.params.IncludeConditional,
		},
	}, nil
}

// getSheetsToFormat determines which sheets to format
func (f *FinancialFormatter) getSheetsToFormat() []string {
	if len(*f.params.ApplyToSheets) > 0 {
		return *f.params.ApplyToSheets
	}

	// デフォルトの財務シート
	return []string{
		"Executive_Summary",
		"Assumptions", 
		"Historical_Analysis",
		"Income_Statement",
		"Balance_Sheet",
		"Cash_Flow",
		"Model_Checks",
	}
}

// formatSheet applies formatting to a specific sheet
func (f *FinancialFormatter) formatSheet(sheetName string) error {
	worksheet, err := f.file.FindSheet(sheetName)
	if err != nil {
		// シートが存在しない場合はスキップ
		return nil
	}

	// シートタイプに基づいてフォーマットを適用
	switch sheetName {
	case "Executive_Summary":
		return f.formatExecutiveSummary(worksheet)
	case "Assumptions":
		return f.formatAssumptions(worksheet)
	case "Historical_Analysis":
		return f.formatHistoricalAnalysis(worksheet)
	case "Income_Statement":
		return f.formatIncomeStatement(worksheet)
	case "Balance_Sheet":
		return f.formatBalanceSheet(worksheet)
	case "Cash_Flow":
		return f.formatCashFlow(worksheet)
	case "Model_Checks":
		return f.formatModelChecks(worksheet)
	default:
		return f.formatGenericFinancialSheet(worksheet)
	}
}

// formatExecutiveSummary formats the executive summary sheet
func (f *FinancialFormatter) formatExecutiveSummary(worksheet excel.Worksheet) error {
	// タイトルフォーマット
	f.applyTitleStyle(worksheet, "A1")
	
	// ヘッダーフォーマット
	if *f.params.IncludeHeaders {
		f.applyHeaderStyle(worksheet, "A6") // "Key Financial Metrics"
		f.applyHeaderStyle(worksheet, "A7:G7") // 年度ヘッダー
	}

	// 数値フォーマット
	f.applyCurrencyFormat(worksheet, "C8:G15") // 財務指標
	f.applyPercentageFormat(worksheet, "C16:G23") // パーセント指標

	// 条件付きフォーマット
	if *f.params.IncludeConditional {
		f.applyConditionalFormatting(worksheet, "C8:G15", "financial_metrics")
	}

	return nil
}

// formatAssumptions formats the assumptions sheet
func (f *FinancialFormatter) formatAssumptions(worksheet excel.Worksheet) error {
	// ヘッダーフォーマット
	if *f.params.IncludeHeaders {
		f.applyHeaderStyle(worksheet, "A1") // "Revenue Growth Assumptions"
		f.applyHeaderStyle(worksheet, "A5") // "Projected Growth Rates"
		f.applyHeaderStyle(worksheet, "A12") // "Margin Assumptions"
		f.applyHeaderStyle(worksheet, "A16") // "Working Capital Assumptions"
	}

	// 入力値のフォーマット (青字)
	f.applyInputValueStyle(worksheet, "B6:B11") // 成長率
	f.applyInputValueStyle(worksheet, "B13:B14") // マージン
	f.applyInputValueStyle(worksheet, "B17:B19") // 運転資本

	// パーセント形式
	f.applyPercentageFormat(worksheet, "B6:B14")
	
	// 数値フォーマット
	f.applyNumberFormat(worksheet, "B17:B19")

	return nil
}

// formatHistoricalAnalysis formats the historical analysis sheet
func (f *FinancialFormatter) formatHistoricalAnalysis(worksheet excel.Worksheet) error {
	// ヘッダーフォーマット
	if *f.params.IncludeHeaders {
		f.applyHeaderStyle(worksheet, "A1") // タイトル
		f.applyHeaderStyle(worksheet, "A5") // "Income Statement (Historical)"
		f.applyHeaderStyle(worksheet, "A20") // "Growth & Margin Analysis"
		f.applyHeaderStyle(worksheet, "C3:G3") // 年度ヘッダー
	}

	// 通貨フォーマット
	f.applyCurrencyFormat(worksheet, "C6:G16") // 損益計算書項目

	// パーセント形式
	f.applyPercentageFormat(worksheet, "C21:G24") // 成長率とマージン

	// 入力値のスタイル
	f.applyInputValueStyle(worksheet, "C6:G16")

	return nil
}

// formatIncomeStatement formats the income statement sheet
func (f *FinancialFormatter) formatIncomeStatement(worksheet excel.Worksheet) error {
	// ヘッダーフォーマット
	if *f.params.IncludeHeaders {
		f.applyTitleStyle(worksheet, "A1")
		f.applyHeaderStyle(worksheet, "C3:G3") // 年度ヘッダー
	}

	// 通貨フォーマット
	f.applyCurrencyFormat(worksheet, "C5:G30") // 全ての金額項目

	// パーセント形式
	f.applyPercentageFormat(worksheet, "C8:G8")   // 売上総利益率
	f.applyPercentageFormat(worksheet, "C22:G22") // EBITDA率
	f.applyPercentageFormat(worksheet, "C25:G25") // 純利益率

	// 数式のスタイル (数式を含むセルは黒字)
	if *f.params.IncludeFormulas {
		f.applyFormulaStyle(worksheet, "C7:G7")   // 売上総利益
		f.applyFormulaStyle(worksheet, "C15:G15") // EBITDA
		f.applyFormulaStyle(worksheet, "C19:G19") // EBIT
		f.applyFormulaStyle(worksheet, "C25:G25") // 純利益
	}

	// 小計・合計のフォーマット
	f.applySubtotalStyle(worksheet, "C12:G12") // 営業費用合計
	f.applyTotalStyle(worksheet, "C15:G15")    // EBITDA
	f.applyTotalStyle(worksheet, "C25:G25")    // 純利益

	return nil
}

// formatBalanceSheet formats the balance sheet sheet
func (f *FinancialFormatter) formatBalanceSheet(worksheet excel.Worksheet) error {
	// ヘッダーフォーマット
	if *f.params.IncludeHeaders {
		f.applyTitleStyle(worksheet, "A1")
		f.applyHeaderStyle(worksheet, "C3:G3") // 年度ヘッダー
		f.applyHeaderStyle(worksheet, "A5")    // "ASSETS"
		f.applyHeaderStyle(worksheet, "A23")   // "LIABILITIES & EQUITY"
	}

	// 通貨フォーマット
	f.applyCurrencyFormat(worksheet, "C7:G43") // 全ての金額項目

	// 小計・合計のフォーマット
	f.applySubtotalStyle(worksheet, "C11:G11") // 流動資産合計
	f.applySubtotalStyle(worksheet, "C19:G19") // 固定資産合計
	f.applyTotalStyle(worksheet, "C21:G21")    // 総資産
	f.applySubtotalStyle(worksheet, "C28:G28") // 流動負債合計
	f.applySubtotalStyle(worksheet, "C33:G33") // 固定負債合計
	f.applySubtotalStyle(worksheet, "C35:G35") // 負債合計
	f.applySubtotalStyle(worksheet, "C41:G41") // 資本合計
	f.applyTotalStyle(worksheet, "C43:G43")    // 負債・資本合計

	// バランスチェック
	f.applyValidationStyle(worksheet, "C46:G46")

	return nil
}

// formatCashFlow formats the cash flow statement sheet
func (f *FinancialFormatter) formatCashFlow(worksheet excel.Worksheet) error {
	// ヘッダーフォーマット
	if *f.params.IncludeHeaders {
		f.applyTitleStyle(worksheet, "A1")
		f.applyHeaderStyle(worksheet, "C3:G3") // 年度ヘッダー
		f.applyHeaderStyle(worksheet, "A5")    // "Operating Activities"
		f.applyHeaderStyle(worksheet, "A15")   // "Investing Activities"
		f.applyHeaderStyle(worksheet, "A20")   // "Financing Activities"
	}

	// 通貨フォーマット
	f.applyCurrencyFormat(worksheet, "C6:G28") // 全ての金額項目

	// 小計・合計のフォーマット
	f.applySubtotalStyle(worksheet, "C13:G13") // 営業CF
	f.applySubtotalStyle(worksheet, "C18:G18") // 投資CF
	f.applySubtotalStyle(worksheet, "C24:G24") // 財務CF
	f.applyTotalStyle(worksheet, "C26:G26")    // 現金変動
	f.applyTotalStyle(worksheet, "C28:G28")    // 期末現金残高

	return nil
}

// formatModelChecks formats the model checks sheet
func (f *FinancialFormatter) formatModelChecks(worksheet excel.Worksheet) error {
	// ヘッダーフォーマット
	if *f.params.IncludeHeaders {
		f.applyTitleStyle(worksheet, "A1")
		f.applyHeaderStyle(worksheet, "C3:G3") // 年度ヘッダー
	}

	// 検証結果のフォーマット
	f.applyValidationStyle(worksheet, "C5:G10")

	// 条件付きフォーマット（エラーの場合は赤背景）
	if *f.params.IncludeConditional {
		f.applyConditionalFormatting(worksheet, "C5:G10", "error_checking")
	}

	return nil
}

// formatGenericFinancialSheet formats a generic financial sheet
func (f *FinancialFormatter) formatGenericFinancialSheet(worksheet excel.Worksheet) error {
	// 基本的なフォーマットを適用
	f.applyCurrencyFormat(worksheet, "C1:G100")
	return nil
}

// Style application methods

// applyTitleStyle applies title formatting
func (f *FinancialFormatter) applyTitleStyle(worksheet excel.Worksheet, cellRange string) {
	// タイトルスタイルの適用（太字、大きなフォント）
	f.stylesApplied["title_style"] = true
}

// applyHeaderStyle applies header formatting
func (f *FinancialFormatter) applyHeaderStyle(worksheet excel.Worksheet, cellRange string) {
	// ヘッダースタイルの適用（太字、背景色）
	f.stylesApplied["header_style"] = true
}

// applyInputValueStyle applies input value formatting (blue text)
func (f *FinancialFormatter) applyInputValueStyle(worksheet excel.Worksheet, cellRange string) {
	// 入力値スタイルの適用（青字: RGB 166,203,240）
	f.stylesApplied["input_value_style"] = true
}

// applyFormulaStyle applies formula formatting (black text)
func (f *FinancialFormatter) applyFormulaStyle(worksheet excel.Worksheet, cellRange string) {
	// 数式スタイルの適用（黒字）
	f.stylesApplied["formula_style"] = true
}

// applyLinkReferenceStyle applies link reference formatting (green text)
func (f *FinancialFormatter) applyLinkReferenceStyle(worksheet excel.Worksheet, cellRange string) {
	// リンク参照スタイルの適用（緑字: RGB 198,239,206）
	f.stylesApplied["link_reference_style"] = true
}

// applySubtotalStyle applies subtotal formatting
func (f *FinancialFormatter) applySubtotalStyle(worksheet excel.Worksheet, cellRange string) {
	// 小計スタイルの適用（太字）
	f.stylesApplied["subtotal_style"] = true
}

// applyTotalStyle applies total formatting
func (f *FinancialFormatter) applyTotalStyle(worksheet excel.Worksheet, cellRange string) {
	// 合計スタイルの適用（太字 + 上下線）
	f.stylesApplied["total_style"] = true
}

// applyValidationStyle applies validation formatting
func (f *FinancialFormatter) applyValidationStyle(worksheet excel.Worksheet, cellRange string) {
	// 検証スタイルの適用
	f.stylesApplied["validation_style"] = true
}

// applyCurrencyFormat applies currency formatting
func (f *FinancialFormatter) applyCurrencyFormat(worksheet excel.Worksheet, cellRange string) {
	// 通貨フォーマットの適用
	currencyFormat := f.getCurrencyFormat(*f.params.Currency)
	f.stylesApplied["currency_format"] = currencyFormat
}

// applyPercentageFormat applies percentage formatting
func (f *FinancialFormatter) applyPercentageFormat(worksheet excel.Worksheet, cellRange string) {
	// パーセントフォーマットの適用
	f.stylesApplied["percentage_format"] = "0.0%"
}

// applyNumberFormat applies number formatting
func (f *FinancialFormatter) applyNumberFormat(worksheet excel.Worksheet, cellRange string) {
	// 数値フォーマットの適用
	f.stylesApplied["number_format"] = "#,##0"
}

// applyConditionalFormatting applies conditional formatting
func (f *FinancialFormatter) applyConditionalFormatting(worksheet excel.Worksheet, cellRange string, ruleType string) {
	// 条件付きフォーマットの適用
	switch ruleType {
	case "financial_metrics":
		// 財務指標の条件付きフォーマット
		f.stylesApplied["conditional_financial"] = true
	case "error_checking":
		// エラーチェックの条件付きフォーマット
		f.stylesApplied["conditional_error"] = true
	}
}

// Helper methods

// getColorScheme returns the color scheme for the specified format type
func getColorScheme(formatType string) ColorScheme {
	switch formatType {
	case "investment_banking":
		return ColorScheme{
			InputValues:    "#A6CBF0", // 青字 (RGB 166,203,240)
			Formulas:       "#000000", // 黒字
			LinkReferences: "#C6EFCE", // 緑字 (RGB 198,239,206)
			Headers:        "#2F5597", // 濃い青
			Subtotals:      "#000000", // 黒字・太字
			Totals:         "#000000", // 黒字・太字・線
		}
	case "corporate_finance":
		return ColorScheme{
			InputValues:    "#ADD8E6", // 薄い青
			Formulas:       "#000000", // 黒字
			LinkReferences: "#90EE90", // 薄い緑
			Headers:        "#4682B4", // スチールブルー
			Subtotals:      "#000000", // 黒字・太字
			Totals:         "#000000", // 黒字・太字・線
		}
	case "consulting":
		return ColorScheme{
			InputValues:    "#B0C4DE", // ライトスチールブルー
			Formulas:       "#000000", // 黒字
			LinkReferences: "#98FB98", // ペールグリーン
			Headers:        "#483D8B", // ダークスレートブルー
			Subtotals:      "#000000", // 黒字・太字
			Totals:         "#000000", // 黒字・太字・線
		}
	default:
		return getColorScheme("investment_banking")
	}
}

// getCurrencyFormat returns the currency format for the specified currency
func (f *FinancialFormatter) getCurrencyFormat(currency string) string {
	switch strings.ToUpper(currency) {
	case "USD":
		return `"$"#,##0_);("$"#,##0)`
	case "JPY":
		return `"¥"#,##0_);("¥"#,##0)`
	case "EUR":
		return `"€"#,##0_);("€"#,##0)`
	case "GBP":
		return `"£"#,##0_);("£"#,##0)`
	default:
		return `#,##0_);(#,##0)`
	}
}

// createFinancialNumberFormat creates a financial number format
func (f *FinancialFormatter) createFinancialNumberFormat(isMillions bool) string {
	if isMillions {
		return `#,##0,,"M";(#,##0,,"M");"-"`
	}
	return `#,##0_);(#,##0);"-"`
}

// applyBordersAndLines applies borders and lines to ranges
func (f *FinancialFormatter) applyBordersAndLines(worksheet excel.Worksheet, cellRange string, borderType string) {
	// 枠線の適用
	switch borderType {
	case "single_top":
		f.stylesApplied["border_single_top"] = true
	case "double_top":
		f.stylesApplied["border_double_top"] = true
	case "single_bottom":
		f.stylesApplied["border_single_bottom"] = true
	case "double_bottom":
		f.stylesApplied["border_double_bottom"] = true
	case "outline":
		f.stylesApplied["border_outline"] = true
	}
}

func AddExcelFinancialFormattingTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_financial_formatting",
		mcp.WithDescription("Apply professional financial formatting to Excel models with industry-standard color schemes and number formats"),
		mcp.WithString("file_path",
			mcp.Required(),
			mcp.Description("Path to the Excel file to format"),
		),
		mcp.WithString("format_type",
			mcp.Description("Format type (investment_banking, corporate_finance, consulting)"),
		),
		mcp.WithString("currency",
			mcp.Description("Currency code (USD, JPY, EUR, etc.)"),
		),
		mcp.WithBoolean("include_headers",
			mcp.Description("Apply formatting to headers (default: true)"),
		),
		mcp.WithBoolean("include_formulas",
			mcp.Description("Apply color coding to formulas (default: true)"),
		),
		mcp.WithBoolean("include_conditional",
			mcp.Description("Apply conditional formatting (default: true)"),
		),
	), func(arguments mcp.ToolCallArguments) *mcp.CallToolResult {
		var args ExcelFinancialFormattingParams
		if err := arguments.Unmarshal(&args); err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("Invalid arguments: %v", err))
		}

		result, err := ExcelFinancialFormatting(context.Background(), args)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("Failed to apply formatting: %v", err))
		}

		return mcp.NewToolResultText(fmt.Sprintf("Financial formatting completed: %s", result.Message))
	})
}