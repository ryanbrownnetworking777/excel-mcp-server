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
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"

	z "github.com/Oudwins/zog"
)

// ExcelFinancialTemplateArguments represents the parameters for creating a financial template
type ExcelFinancialTemplateArguments struct {
	FilePath        string `zog:"filePath"`
	CompanyName     string `zog:"companyName"`
	Currency        string `zog:"currency"`
	FiscalYearEnd   string `zog:"fiscalYearEnd"`
	IndustryType    string `zog:"industryType"`
	ProjectionYears int    `zog:"projectionYears"`
}

// ExcelFinancialTemplateResult represents the result of creating a financial template
type ExcelFinancialTemplateResult struct {
	Message    string                 `json:"message"`
	FileInfo   map[string]interface{} `json:"file_info"`
	ModelStats map[string]interface{} `json:"model_stats"`
}

var excelFinancialTemplateArgumentsSchema = z.Struct(z.Schema{
	"filePath":        z.String().Test(AbsolutePathTest()).Required(),
	"companyName":     z.String().Min(1).Required(),
	"currency":        z.String().Min(1).Required(),
	"fiscalYearEnd":   z.String().Min(1).Required(),
	"industryType":    z.String().Min(1).Required(),
	"projectionYears": z.Int().Min(1).Max(10).Default(5),
})

// ExcelFinancialTemplate creates a comprehensive three-statement financial model template
func ExcelFinancialTemplate(args ExcelFinancialTemplateArguments) (*mcp.CallToolResult, error) {
	// Excelファイルを新規作成
	excelFile, closeFn, err := excel.CreateNewFile(args.FilePath)
	if err != nil {
		return nil, fmt.Errorf("failed to create Excel file: %w", err)
	}
	defer closeFn()

	// 三表統合テンプレートの作成
	err = createFinancialModelTemplate(excelFile, args)
	if err != nil {
		return nil, fmt.Errorf("failed to create financial template: %w", err)
	}

	// ファイルを保存
	err = excelFile.Save()
	if err != nil {
		return nil, fmt.Errorf("failed to save Excel file: %w", err)
	}

	// 結果を返す
	message := fmt.Sprintf("Financial model template created successfully for %s", args.CompanyName)
	return mcp.NewCallToolResult(mcp.TextContent(message)), nil
}

// createFinancialModelTemplate creates the complete financial model structure
func createFinancialModelTemplate(excelFile excel.Excel, args ExcelFinancialTemplateArguments) error {
	// 基本情報
	projectionYears := args.ProjectionYears
	
	// 各シートを作成
	sheets := []string{
		"Executive_Summary",
		"Assumptions", 
		"Historical_Analysis",
		"Income_Statement",
		"Balance_Sheet",
		"Cash_Flow",
		"Model_Checks",
	}

	for _, sheetName := range sheets {
		err := excelFile.CreateNewSheet(sheetName)
		if err != nil {
			return fmt.Errorf("failed to create sheet %s: %w", sheetName, err)
		}
	}

	// デフォルトシートを削除
	err := excelFile.DeleteSheet("Sheet1")
	if err != nil {
		// Sheet1が存在しない場合は無視
	}

	// 各シートのテンプレートを作成
	if err := createExecutiveSummarySheet(excelFile, args); err != nil {
		return err
	}
	if err := createAssumptionsSheet(excelFile, args); err != nil {
		return err
	}
	if err := createHistoricalAnalysisSheet(excelFile, args); err != nil {
		return err
	}
	if err := createIncomeStatementSheet(excelFile, args); err != nil {
		return err
	}
	if err := createBalanceSheetSheet(excelFile, args); err != nil {
		return err
	}
	if err := createCashFlowSheet(excelFile, args); err != nil {
		return err
	}
	if err := createModelChecksSheet(excelFile, args); err != nil {
		return err
	}

	return nil
}

// createExecutiveSummarySheet creates the executive summary dashboard
func createExecutiveSummarySheet(excelFile excel.Excel, args ExcelFinancialTemplateArguments) error {
	worksheet, err := excelFile.FindSheet("Executive_Summary")
	if err != nil {
		return err
	}

	// ヘッダー情報
	worksheet.SetValue("A1", fmt.Sprintf("%s - Financial Model", args.CompanyName))
	worksheet.SetValue("A2", fmt.Sprintf("Industry: %s", args.IndustryType))
	worksheet.SetValue("A3", fmt.Sprintf("Currency: %s", args.Currency))
	worksheet.SetValue("A4", fmt.Sprintf("Fiscal Year End: %s", args.FiscalYearEnd))

	// 主要財務指標のテーブル
	worksheet.SetValue("A6", "Key Financial Metrics")
	
	// 年度ヘッダー
	startYear := 2024
	for i := 0; i < args.ProjectionYears; i++ {
		col := rune('C' + i)
		worksheet.SetValue(fmt.Sprintf("%c7", col), fmt.Sprintf("FY%d", startYear+i))
	}

	// 主要指標
	metrics := []string{
		"Revenue",
		"Gross Profit",
		"EBITDA", 
		"Net Income",
		"Revenue Growth %",
		"Gross Margin %",
		"EBITDA Margin %",
		"Net Margin %",
	}

	for i, metric := range metrics {
		worksheet.SetValue(fmt.Sprintf("A%d", 8+i), metric)
	}

	return nil
}

// createAssumptionsSheet creates the assumptions sheet
func createAssumptionsSheet(excelFile excel.Excel, params ExcelFinancialTemplateParams) error {
	worksheet, err := excelFile.FindSheet("Assumptions")
	if err != nil {
		return err
	}

	// 収益成長率の前提条件
	worksheet.SetValue("A1", "Revenue Growth Assumptions")
	worksheet.SetValue("A3", "Historical Growth Rate (5-year CAGR)")
	worksheet.SetValue("B3", "Enter historical data")
	
	worksheet.SetValue("A5", "Projected Growth Rates")
	
	// 業界タイプによる初期値設定
	growthRates := getIndustryGrowthRates(*params.IndustryType)
	
	for i := 0; i < *params.ProjectionYears; i++ {
		year := 2024 + i
		worksheet.SetValue(fmt.Sprintf("A%d", 6+i), fmt.Sprintf("FY%d Growth Rate", year))
		worksheet.SetValue(fmt.Sprintf("B%d", 6+i), growthRates[i])
	}

	// マージンの前提条件
	worksheet.SetValue("A12", "Margin Assumptions")
	worksheet.SetValue("A13", "Gross Margin %")
	worksheet.SetValue("B13", getIndustryMargin(*params.IndustryType, "gross"))
	worksheet.SetValue("A14", "EBITDA Margin %")
	worksheet.SetValue("B14", getIndustryMargin(*params.IndustryType, "ebitda"))

	// 運転資本の前提条件
	worksheet.SetValue("A16", "Working Capital Assumptions")
	worksheet.SetValue("A17", "Days Sales Outstanding (DSO)")
	worksheet.SetValue("B17", "45")
	worksheet.SetValue("A18", "Days Inventory Outstanding (DIO)")
	worksheet.SetValue("B18", "60")
	worksheet.SetValue("A19", "Days Payable Outstanding (DPO)")
	worksheet.SetValue("B19", "30")

	return nil
}

// createHistoricalAnalysisSheet creates the historical analysis sheet
func createHistoricalAnalysisSheet(excelFile excel.Excel, params ExcelFinancialTemplateParams) error {
	worksheet, err := excelFile.FindSheet("Historical_Analysis")
	if err != nil {
		return err
	}

	worksheet.SetValue("A1", "Historical Financial Analysis")
	worksheet.SetValue("A2", "Enter historical data for the past 5 years")

	// 年度ヘッダー
	startYear := 2019
	for i := 0; i < 5; i++ {
		col := rune('C' + i)
		worksheet.SetValue(fmt.Sprintf("%c3", col), fmt.Sprintf("FY%d", startYear+i))
	}

	// 損益計算書の歴史的データ
	incomeItems := []string{
		"Revenue",
		"Cost of Goods Sold",
		"Gross Profit",
		"Operating Expenses",
		"EBITDA",
		"Depreciation & Amortization",
		"EBIT",
		"Interest Expense",
		"Pre-tax Income",
		"Tax Expense",
		"Net Income",
	}

	worksheet.SetValue("A5", "Income Statement (Historical)")
	for i, item := range incomeItems {
		worksheet.SetValue(fmt.Sprintf("A%d", 6+i), item)
	}

	// 成長率とマージン分析
	worksheet.SetValue("A20", "Growth & Margin Analysis")
	worksheet.SetValue("A21", "Revenue Growth %")
	worksheet.SetValue("A22", "Gross Margin %")
	worksheet.SetValue("A23", "EBITDA Margin %")
	worksheet.SetValue("A24", "Net Margin %")

	return nil
}

// createIncomeStatementSheet creates the income statement projection
func createIncomeStatementSheet(excelFile excel.Excel, params ExcelFinancialTemplateParams) error {
	worksheet, err := excelFile.FindSheet("Income_Statement")
	if err != nil {
		return err
	}

	worksheet.SetValue("A1", fmt.Sprintf("%s - Income Statement Projection", *params.CompanyName))
	worksheet.SetValue("A2", fmt.Sprintf("(%s millions)", *params.Currency))

	// 年度ヘッダー
	startYear := 2024
	for i := 0; i < args.ProjectionYears; i++ {
		col := rune('C' + i)
		worksheet.SetValue(fmt.Sprintf("%c3", col), fmt.Sprintf("FY%d", startYear+i))
	}

	// 損益計算書の項目
	incomeItems := []string{
		"Revenue",
		"Cost of Goods Sold",
		"Gross Profit",
		"% of Revenue",
		"", // 空行
		"Operating Expenses",
		"Sales & Marketing",
		"General & Administrative",
		"Research & Development",
		"Total Operating Expenses",
		"", // 空行
		"EBITDA",
		"% of Revenue",
		"", // 空行
		"Depreciation & Amortization",
		"EBIT",
		"% of Revenue",
		"", // 空行
		"Interest Expense",
		"Other Income (Expense)",
		"Pre-tax Income",
		"", // 空行
		"Tax Rate",
		"Tax Expense",
		"Net Income",
		"% of Revenue",
	}

	for i, item := range incomeItems {
		if item != "" {
			worksheet.SetValue(fmt.Sprintf("A%d", 5+i), item)
		}
	}

	// 基本的な数式を設定
	err = addIncomeStatementFormulas(worksheet, *params.ProjectionYears)
	if err != nil {
		return err
	}

	return nil
}

// createBalanceSheetSheet creates the balance sheet projection
func createBalanceSheetSheet(excelFile excel.Excel, params ExcelFinancialTemplateParams) error {
	worksheet, err := excelFile.FindSheet("Balance_Sheet")
	if err != nil {
		return err
	}

	worksheet.SetValue("A1", fmt.Sprintf("%s - Balance Sheet Projection", *params.CompanyName))
	worksheet.SetValue("A2", fmt.Sprintf("(%s millions)", *params.Currency))

	// 年度ヘッダー
	startYear := 2024
	for i := 0; i < args.ProjectionYears; i++ {
		col := rune('C' + i)
		worksheet.SetValue(fmt.Sprintf("%c3", col), fmt.Sprintf("FY%d", startYear+i))
	}

	// 資産の部
	worksheet.SetValue("A5", "ASSETS")
	worksheet.SetValue("A6", "Current Assets")
	worksheet.SetValue("A7", "Cash & Cash Equivalents")
	worksheet.SetValue("A8", "Accounts Receivable")
	worksheet.SetValue("A9", "Inventory")
	worksheet.SetValue("A10", "Other Current Assets")
	worksheet.SetValue("A11", "Total Current Assets")

	worksheet.SetValue("A13", "Non-Current Assets")
	worksheet.SetValue("A14", "Property, Plant & Equipment")
	worksheet.SetValue("A15", "Accumulated Depreciation")
	worksheet.SetValue("A16", "Net PP&E")
	worksheet.SetValue("A17", "Intangible Assets")
	worksheet.SetValue("A18", "Other Non-Current Assets")
	worksheet.SetValue("A19", "Total Non-Current Assets")

	worksheet.SetValue("A21", "TOTAL ASSETS")

	// 負債・資本の部
	worksheet.SetValue("A23", "LIABILITIES & EQUITY")
	worksheet.SetValue("A24", "Current Liabilities")
	worksheet.SetValue("A25", "Accounts Payable")
	worksheet.SetValue("A26", "Short-term Debt")
	worksheet.SetValue("A27", "Accrued Expenses")
	worksheet.SetValue("A28", "Total Current Liabilities")

	worksheet.SetValue("A30", "Non-Current Liabilities")
	worksheet.SetValue("A31", "Long-term Debt")
	worksheet.SetValue("A32", "Other Non-Current Liabilities")
	worksheet.SetValue("A33", "Total Non-Current Liabilities")

	worksheet.SetValue("A35", "Total Liabilities")

	worksheet.SetValue("A37", "Shareholders' Equity")
	worksheet.SetValue("A38", "Share Capital")
	worksheet.SetValue("A39", "Retained Earnings")
	worksheet.SetValue("A40", "Other Equity")
	worksheet.SetValue("A41", "Total Shareholders' Equity")

	worksheet.SetValue("A43", "TOTAL LIABILITIES & EQUITY")

	// バランスチェック
	worksheet.SetValue("A45", "Balance Check")
	worksheet.SetValue("A46", "Difference (Should be 0)")

	return nil
}

// createCashFlowSheet creates the cash flow statement projection
func createCashFlowSheet(excelFile excel.Excel, params ExcelFinancialTemplateParams) error {
	worksheet, err := excelFile.FindSheet("Cash_Flow")
	if err != nil {
		return err
	}

	worksheet.SetValue("A1", fmt.Sprintf("%s - Cash Flow Statement Projection", *params.CompanyName))
	worksheet.SetValue("A2", fmt.Sprintf("(%s millions)", *params.Currency))

	// 年度ヘッダー
	startYear := 2024
	for i := 0; i < args.ProjectionYears; i++ {
		col := rune('C' + i)
		worksheet.SetValue(fmt.Sprintf("%c3", col), fmt.Sprintf("FY%d", startYear+i))
	}

	// 営業活動によるキャッシュフロー
	worksheet.SetValue("A5", "Operating Activities")
	worksheet.SetValue("A6", "Net Income")
	worksheet.SetValue("A7", "Adjustments:")
	worksheet.SetValue("A8", "Depreciation & Amortization")
	worksheet.SetValue("A9", "Changes in Working Capital:")
	worksheet.SetValue("A10", "Change in Accounts Receivable")
	worksheet.SetValue("A11", "Change in Inventory")
	worksheet.SetValue("A12", "Change in Accounts Payable")
	worksheet.SetValue("A13", "Net Cash from Operating Activities")

	// 投資活動によるキャッシュフロー
	worksheet.SetValue("A15", "Investing Activities")
	worksheet.SetValue("A16", "Capital Expenditures")
	worksheet.SetValue("A17", "Other Investing Activities")
	worksheet.SetValue("A18", "Net Cash from Investing Activities")

	// 財務活動によるキャッシュフロー
	worksheet.SetValue("A20", "Financing Activities")
	worksheet.SetValue("A21", "Net Borrowings")
	worksheet.SetValue("A22", "Dividends Paid")
	worksheet.SetValue("A23", "Other Financing Activities")
	worksheet.SetValue("A24", "Net Cash from Financing Activities")

	// 現金の変動
	worksheet.SetValue("A26", "Net Change in Cash")
	worksheet.SetValue("A27", "Beginning Cash Balance")
	worksheet.SetValue("A28", "Ending Cash Balance")

	return nil
}

// createModelChecksSheet creates the model validation sheet
func createModelChecksSheet(excelFile excel.Excel, params ExcelFinancialTemplateParams) error {
	worksheet, err := excelFile.FindSheet("Model_Checks")
	if err != nil {
		return err
	}

	worksheet.SetValue("A1", "Model Integrity Checks")
	worksheet.SetValue("A2", "All checks should equal 0 or show OK")

	// 年度ヘッダー
	startYear := 2024
	for i := 0; i < args.ProjectionYears; i++ {
		col := rune('C' + i)
		worksheet.SetValue(fmt.Sprintf("%c3", col), fmt.Sprintf("FY%d", startYear+i))
	}

	// 整合性チェック項目
	checks := []string{
		"Balance Sheet Check",
		"Cash Flow Check",
		"Working Capital Check",
		"Debt Interest Check",
		"Tax Rate Check",
		"Circular Reference Check",
	}

	for i, check := range checks {
		worksheet.SetValue(fmt.Sprintf("A%d", 5+i), check)
	}

	return nil
}

// addIncomeStatementFormulas adds basic formulas to the income statement
func addIncomeStatementFormulas(worksheet excel.Worksheet, projectionYears int) error {
	// 基本的な数式の例
	for i := 0; i < projectionYears; i++ {
		col := rune('C' + i)
		colStr := string(col)
		
		// Gross Profit = Revenue - COGS
		worksheet.SetFormula(fmt.Sprintf("%s7", colStr), fmt.Sprintf("=%s5-%s6", colStr, colStr))
		
		// Gross Margin %
		worksheet.SetFormula(fmt.Sprintf("%s8", colStr), fmt.Sprintf("=%s7/%s5", colStr, colStr))
	}

	return nil
}

// getIndustryGrowthRates returns industry-specific growth rates
func getIndustryGrowthRates(industryType string) []float64 {
	switch strings.ToLower(industryType) {
	case "technology":
		return []float64{0.15, 0.12, 0.10, 0.08, 0.06}
	case "manufacturing":
		return []float64{0.05, 0.04, 0.03, 0.03, 0.03}
	case "retail":
		return []float64{0.06, 0.05, 0.04, 0.04, 0.03}
	case "financial":
		return []float64{0.08, 0.07, 0.06, 0.05, 0.05}
	default:
		return []float64{0.05, 0.05, 0.04, 0.04, 0.03}
	}
}

// getIndustryMargin returns industry-specific margin
func getIndustryMargin(industryType string, marginType string) float64 {
	margins := map[string]map[string]float64{
		"technology": {"gross": 0.70, "ebitda": 0.25},
		"manufacturing": {"gross": 0.35, "ebitda": 0.12},
		"retail": {"gross": 0.40, "ebitda": 0.08},
		"financial": {"gross": 0.60, "ebitda": 0.20},
	}

	if industry, exists := margins[strings.ToLower(industryType)]; exists {
		if margin, exists := industry[marginType]; exists {
			return margin
		}
	}
	
	// デフォルト値
	if marginType == "gross" {
		return 0.40
	}
	return 0.15
}

func AddExcelFinancialTemplateTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_financial_template",
		mcp.WithDescription("Create a comprehensive three-statement financial model template with industry-specific assumptions"),
		mcp.WithString("filePath",
			mcp.Required(),
			mcp.Description("Path to the Excel file to create"),
		),
		mcp.WithString("companyName",
			mcp.Required(),
			mcp.Description("Company name for the financial model"),
		),
		mcp.WithString("currency",
			mcp.Required(),
			mcp.Description("Currency code (e.g., USD, JPY)"),
		),
		mcp.WithString("fiscalYearEnd",
			mcp.Required(),
			mcp.Description("Fiscal year end month (e.g., December, March)"),
		),
		mcp.WithString("industryType",
			mcp.Required(),
			mcp.Description("Industry type (technology, manufacturing, retail, financial)"),
		),
		mcp.WithNumber("projectionYears",
			mcp.Description("Number of years to project (default: 5)"),
		),
	), handleFinancialTemplate)
}

func handleFinancialTemplate(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelFinancialTemplateArguments{}
	issues := excelFinancialTemplateArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}

	return ExcelFinancialTemplate(args)
}