package tools

import (
	"context"
	"fmt"
	"math"
	"sort"
	"strconv"
	"strings"

	"github.com/negokaz/excel-mcp-server/internal/excel"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"

	z "github.com/Oudwins/zog"
)

// ExcelHistoricalAnalysisParams represents the parameters for historical analysis
type ExcelHistoricalAnalysisParams struct {
	FilePath         *string `json:"file_path" doc:"Path to the Excel file containing historical data"`
	DataSheet        *string `json:"data_sheet" doc:"Name of the sheet containing historical data (default: Historical_Analysis)"`
	StartYear        *int    `json:"start_year" doc:"Start year for analysis (default: 2019)"`
	EndYear          *int    `json:"end_year" doc:"End year for analysis (default: 2023)"`
	AnalysisType     *string `json:"analysis_type" doc:"Type of analysis (comprehensive, growth, margins, efficiency)"`
	CompanyName      *string `json:"company_name" doc:"Company name for analysis context"`
	IndustryType     *string `json:"industry_type" doc:"Industry type for benchmark comparison"`
	OutputSheet      *string `json:"output_sheet" doc:"Name of output sheet (default: Analysis_Results)"`
	IncludeCharts    *bool   `json:"include_charts" doc:"Include charts in analysis (default: true)"`
	IncludeForecasts *bool   `json:"include_forecasts" doc:"Include forecast suggestions (default: true)"`
}

// ExcelHistoricalAnalysisResult represents the result of historical analysis
type ExcelHistoricalAnalysisResult struct {
	Message            string                 `json:"message"`
	AnalysisSummary    map[string]interface{} `json:"analysis_summary"`
	GrowthAnalysis     GrowthAnalysis         `json:"growth_analysis"`
	MarginAnalysis     MarginAnalysis         `json:"margin_analysis"`
	EfficiencyAnalysis EfficiencyAnalysis     `json:"efficiency_analysis"`
	QualityAssessment  QualityAssessment      `json:"quality_assessment"`
	ForecastSuggestions []ForecastSuggestion  `json:"forecast_suggestions"`
	BenchmarkComparison map[string]interface{} `json:"benchmark_comparison"`
}

// GrowthAnalysis represents growth rate analysis
type GrowthAnalysis struct {
	RevenueGrowth     GrowthMetrics `json:"revenue_growth"`
	ProfitGrowth      GrowthMetrics `json:"profit_growth"`
	AssetGrowth       GrowthMetrics `json:"asset_growth"`
	OverallTrend      string        `json:"overall_trend"`
	SeasonalityIndex  float64       `json:"seasonality_index"`
	VolatilityScore   float64       `json:"volatility_score"`
}

// MarginAnalysis represents margin analysis
type MarginAnalysis struct {
	GrossMargin    MarginMetrics `json:"gross_margin"`
	EBITDAMargin   MarginMetrics `json:"ebitda_margin"`
	NetMargin      MarginMetrics `json:"net_margin"`
	MarginTrend    string        `json:"margin_trend"`
	MarginStability float64      `json:"margin_stability"`
}

// EfficiencyAnalysis represents efficiency analysis
type EfficiencyAnalysis struct {
	WorkingCapitalEfficiency WorkingCapitalMetrics `json:"working_capital_efficiency"`
	AssetTurnover           EfficiencyMetrics     `json:"asset_turnover"`
	CapitalEfficiency       EfficiencyMetrics     `json:"capital_efficiency"`
	OverallEfficiency       float64               `json:"overall_efficiency"`
}

// GrowthMetrics represents growth metrics
type GrowthMetrics struct {
	CAGR3Year    float64   `json:"cagr_3_year"`
	CAGR5Year    float64   `json:"cagr_5_year"`
	YearlyGrowth []float64 `json:"yearly_growth"`
	AverageGrowth float64  `json:"average_growth"`
	TrendDirection string  `json:"trend_direction"`
}

// MarginMetrics represents margin metrics
type MarginMetrics struct {
	Current     float64   `json:"current"`
	Historical  []float64 `json:"historical"`
	Average     float64   `json:"average"`
	Trend       string    `json:"trend"`
	Volatility  float64   `json:"volatility"`
}

// EfficiencyMetrics represents efficiency metrics
type EfficiencyMetrics struct {
	Current    float64   `json:"current"`
	Historical []float64 `json:"historical"`
	Average    float64   `json:"average"`
	Trend      string    `json:"trend"`
	Improvement float64  `json:"improvement"`
}

// WorkingCapitalMetrics represents working capital metrics
type WorkingCapitalMetrics struct {
	DSO          EfficiencyMetrics `json:"dso"`
	DIO          EfficiencyMetrics `json:"dio"`
	DPO          EfficiencyMetrics `json:"dpo"`
	CashCycle    EfficiencyMetrics `json:"cash_cycle"`
	WCIntensity  EfficiencyMetrics `json:"wc_intensity"`
}

// QualityAssessment represents data quality assessment
type QualityAssessment struct {
	DataCompleteness  float64 `json:"data_completeness"`
	DataConsistency   float64 `json:"data_consistency"`
	OutlierCount      int     `json:"outlier_count"`
	ReliabilityScore  float64 `json:"reliability_score"`
	RecommendedActions []string `json:"recommended_actions"`
}

// ForecastSuggestion represents forecast suggestions
type ForecastSuggestion struct {
	Category      string  `json:"category"`
	Metric        string  `json:"metric"`
	SuggestedValue float64 `json:"suggested_value"`
	Confidence    float64 `json:"confidence"`
	Rationale     string  `json:"rationale"`
	RiskFactors   []string `json:"risk_factors"`
}

// ExcelHistoricalAnalysis performs comprehensive historical financial analysis
func ExcelHistoricalAnalysis(ctx context.Context, params ExcelHistoricalAnalysisParams) (*ExcelHistoricalAnalysisResult, error) {
	// バリデーション
	schema := z.Struct(z.Schema{
		"file_path":         z.String().Test(AbsolutePathTest()),
		"data_sheet":        z.String().Default("Historical_Analysis"),
		"start_year":        z.Int().Min(2000).Max(2030).Default(2019),
		"end_year":          z.Int().Min(2000).Max(2030).Default(2023),
		"analysis_type":     z.String().OneOf("comprehensive", "growth", "margins", "efficiency").Default("comprehensive"),
		"company_name":      z.String().Default("Company"),
		"industry_type":     z.String().Default("general"),
		"output_sheet":      z.String().Default("Analysis_Results"),
		"include_charts":    z.Bool().Default(true),
		"include_forecasts": z.Bool().Default(true),
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

	// アナライザーの初期化
	analyzer := &HistoricalAnalyzer{
		file:   excelFile,
		params: validatedParams,
	}

	// 歴史的データの読み込み
	historicalData, err := analyzer.loadHistoricalData()
	if err != nil {
		return nil, fmt.Errorf("failed to load historical data: %w", err)
	}

	// 分析の実行
	result, err := analyzer.performAnalysis(historicalData)
	if err != nil {
		return nil, fmt.Errorf("analysis failed: %w", err)
	}

	// 結果の出力
	err = analyzer.outputResults(result)
	if err != nil {
		return nil, fmt.Errorf("failed to output results: %w", err)
	}

	// ファイルを保存
	err = excelFile.Save()
	if err != nil {
		return nil, fmt.Errorf("failed to save Excel file: %w", err)
	}

	return result, nil
}

// HistoricalAnalyzer handles the analysis logic
type HistoricalAnalyzer struct {
	file   excel.Excel
	params ExcelHistoricalAnalysisParams
}

// HistoricalData represents historical financial data
type HistoricalData struct {
	Years       []int
	Revenue     []float64
	GrossProfit []float64
	EBITDA      []float64
	NetIncome   []float64
	TotalAssets []float64
	Receivables []float64
	Inventory   []float64
	Payables    []float64
	COGS        []float64
}

// loadHistoricalData loads historical data from the Excel file
func (h *HistoricalAnalyzer) loadHistoricalData() (*HistoricalData, error) {
	worksheet, err := h.file.FindSheet(*h.params.DataSheet)
	if err != nil {
		return nil, fmt.Errorf("data sheet not found: %w", err)
	}

	data := &HistoricalData{}
	
	// 年度の読み込み
	yearCount := *h.params.EndYear - *h.params.StartYear + 1
	for i := 0; i < yearCount; i++ {
		data.Years = append(data.Years, *h.params.StartYear + i)
	}

	// データの読み込み（行番号は実際のシート構造に合わせて調整）
	data.Revenue = h.loadDataRow(worksheet, 6, yearCount)      // 売上高
	data.GrossProfit = h.loadDataRow(worksheet, 8, yearCount) // 売上総利益
	data.EBITDA = h.loadDataRow(worksheet, 12, yearCount)     // EBITDA
	data.NetIncome = h.loadDataRow(worksheet, 16, yearCount)  // 純利益
	data.TotalAssets = h.loadDataRow(worksheet, 20, yearCount) // 総資産
	data.Receivables = h.loadDataRow(worksheet, 24, yearCount) // 売掛金
	data.Inventory = h.loadDataRow(worksheet, 25, yearCount)   // 棚卸資産
	data.Payables = h.loadDataRow(worksheet, 26, yearCount)    // 買掛金
	data.COGS = h.loadDataRow(worksheet, 7, yearCount)        // 売上原価

	return data, nil
}

// loadDataRow loads a row of data from the worksheet
func (h *HistoricalAnalyzer) loadDataRow(worksheet excel.Worksheet, row int, count int) []float64 {
	var values []float64
	
	for i := 0; i < count; i++ {
		col := rune('C' + i)
		cellValue := h.getNumericValue(worksheet, fmt.Sprintf("%c%d", col, row))
		values = append(values, cellValue)
	}
	
	return values
}

// performAnalysis performs the comprehensive analysis
func (h *HistoricalAnalyzer) performAnalysis(data *HistoricalData) (*ExcelHistoricalAnalysisResult, error) {
	// 成長率分析
	growthAnalysis := h.analyzeGrowth(data)
	
	// マージン分析
	marginAnalysis := h.analyzeMargins(data)
	
	// 効率性分析
	efficiencyAnalysis := h.analyzeEfficiency(data)
	
	// データ品質評価
	qualityAssessment := h.assessDataQuality(data)
	
	// 予測提案
	var forecastSuggestions []ForecastSuggestion
	if *h.params.IncludeForecasts {
		forecastSuggestions = h.generateForecastSuggestions(data, growthAnalysis, marginAnalysis)
	}
	
	// ベンチマーク比較
	benchmarkComparison := h.compareToBenchmarks(data, marginAnalysis)

	// 分析サマリー
	summary := map[string]interface{}{
		"company_name":      *h.params.CompanyName,
		"analysis_period":   fmt.Sprintf("%d-%d", *h.params.StartYear, *h.params.EndYear),
		"data_points":       len(data.Years),
		"revenue_cagr":      growthAnalysis.RevenueGrowth.CAGR5Year,
		"avg_gross_margin":  marginAnalysis.GrossMargin.Average,
		"data_quality":      qualityAssessment.ReliabilityScore,
		"forecast_confidence": h.calculateOverallConfidence(forecastSuggestions),
	}

	return &ExcelHistoricalAnalysisResult{
		Message:             fmt.Sprintf("Historical analysis completed for %s (%d-%d)", *h.params.CompanyName, *h.params.StartYear, *h.params.EndYear),
		AnalysisSummary:     summary,
		GrowthAnalysis:      growthAnalysis,
		MarginAnalysis:      marginAnalysis,
		EfficiencyAnalysis:  efficiencyAnalysis,
		QualityAssessment:   qualityAssessment,
		ForecastSuggestions: forecastSuggestions,
		BenchmarkComparison: benchmarkComparison,
	}, nil
}

// analyzeGrowth analyzes growth metrics
func (h *HistoricalAnalyzer) analyzeGrowth(data *HistoricalData) GrowthAnalysis {
	// 売上成長率分析
	revenueGrowth := h.calculateGrowthMetrics(data.Revenue)
	
	// 利益成長率分析
	profitGrowth := h.calculateGrowthMetrics(data.NetIncome)
	
	// 資産成長率分析
	assetGrowth := h.calculateGrowthMetrics(data.TotalAssets)
	
	// 全体的なトレンド
	overallTrend := h.determineOverallTrend([]GrowthMetrics{revenueGrowth, profitGrowth, assetGrowth})
	
	// 季節性指数
	seasonalityIndex := h.calculateSeasonality(data.Revenue)
	
	// ボラティリティスコア
	volatilityScore := h.calculateVolatility(data.Revenue)

	return GrowthAnalysis{
		RevenueGrowth:    revenueGrowth,
		ProfitGrowth:     profitGrowth,
		AssetGrowth:      assetGrowth,
		OverallTrend:     overallTrend,
		SeasonalityIndex: seasonalityIndex,
		VolatilityScore:  volatilityScore,
	}
}

// analyzeMargins analyzes margin metrics
func (h *HistoricalAnalyzer) analyzeMargins(data *HistoricalData) MarginAnalysis {
	// 売上総利益率
	grossMargin := h.calculateMarginMetrics(data.GrossProfit, data.Revenue)
	
	// EBITDA利益率
	ebitdaMargin := h.calculateMarginMetrics(data.EBITDA, data.Revenue)
	
	// 純利益率
	netMargin := h.calculateMarginMetrics(data.NetIncome, data.Revenue)
	
	// マージントレンド
	marginTrend := h.determineMarginTrend(grossMargin, ebitdaMargin, netMargin)
	
	// マージン安定性
	marginStability := h.calculateMarginStability(grossMargin, ebitdaMargin, netMargin)

	return MarginAnalysis{
		GrossMargin:     grossMargin,
		EBITDAMargin:    ebitdaMargin,
		NetMargin:       netMargin,
		MarginTrend:     marginTrend,
		MarginStability: marginStability,
	}
}

// analyzeEfficiency analyzes efficiency metrics
func (h *HistoricalAnalyzer) analyzeEfficiency(data *HistoricalData) EfficiencyAnalysis {
	// 運転資本効率
	workingCapitalEfficiency := h.calculateWorkingCapitalEfficiency(data)
	
	// 資産回転率
	assetTurnover := h.calculateEfficiencyMetrics(data.Revenue, data.TotalAssets)
	
	// 資本効率（簡略化）
	capitalEfficiency := assetTurnover // 実際はROE等を使用
	
	// 全体的な効率性
	overallEfficiency := h.calculateOverallEfficiency(workingCapitalEfficiency, assetTurnover)

	return EfficiencyAnalysis{
		WorkingCapitalEfficiency: workingCapitalEfficiency,
		AssetTurnover:           assetTurnover,
		CapitalEfficiency:       capitalEfficiency,
		OverallEfficiency:       overallEfficiency,
	}
}

// calculateGrowthMetrics calculates growth metrics for a data series
func (h *HistoricalAnalyzer) calculateGrowthMetrics(values []float64) GrowthMetrics {
	var yearlyGrowth []float64
	var validGrowths []float64
	
	// 年間成長率の計算
	for i := 1; i < len(values); i++ {
		if values[i-1] != 0 {
			growth := (values[i] - values[i-1]) / values[i-1]
			yearlyGrowth = append(yearlyGrowth, growth)
			validGrowths = append(validGrowths, growth)
		}
	}
	
	// 平均成長率
	avgGrowth := h.calculateAverage(validGrowths)
	
	// CAGR計算
	cagr3Year := h.calculateCAGR(values, 3)
	cagr5Year := h.calculateCAGR(values, 5)
	
	// トレンド方向
	trendDirection := h.determineTrendDirection(yearlyGrowth)

	return GrowthMetrics{
		CAGR3Year:      cagr3Year,
		CAGR5Year:      cagr5Year,
		YearlyGrowth:   yearlyGrowth,
		AverageGrowth:  avgGrowth,
		TrendDirection: trendDirection,
	}
}

// calculateMarginMetrics calculates margin metrics
func (h *HistoricalAnalyzer) calculateMarginMetrics(numerator, denominator []float64) MarginMetrics {
	var margins []float64
	
	for i := 0; i < len(numerator) && i < len(denominator); i++ {
		if denominator[i] != 0 {
			margin := numerator[i] / denominator[i]
			margins = append(margins, margin)
		}
	}
	
	current := 0.0
	if len(margins) > 0 {
		current = margins[len(margins)-1]
	}
	
	average := h.calculateAverage(margins)
	trend := h.determineTrendDirection(margins)
	volatility := h.calculateVolatility(margins)

	return MarginMetrics{
		Current:    current,
		Historical: margins,
		Average:    average,
		Trend:      trend,
		Volatility: volatility,
	}
}

// calculateEfficiencyMetrics calculates efficiency metrics
func (h *HistoricalAnalyzer) calculateEfficiencyMetrics(numerator, denominator []float64) EfficiencyMetrics {
	var ratios []float64
	
	for i := 0; i < len(numerator) && i < len(denominator); i++ {
		if denominator[i] != 0 {
			ratio := numerator[i] / denominator[i]
			ratios = append(ratios, ratio)
		}
	}
	
	current := 0.0
	if len(ratios) > 0 {
		current = ratios[len(ratios)-1]
	}
	
	average := h.calculateAverage(ratios)
	trend := h.determineTrendDirection(ratios)
	
	// 改善度合い
	improvement := 0.0
	if len(ratios) > 1 {
		improvement = (ratios[len(ratios)-1] - ratios[0]) / ratios[0]
	}

	return EfficiencyMetrics{
		Current:     current,
		Historical:  ratios,
		Average:     average,
		Trend:       trend,
		Improvement: improvement,
	}
}

// calculateWorkingCapitalEfficiency calculates working capital efficiency
func (h *HistoricalAnalyzer) calculateWorkingCapitalEfficiency(data *HistoricalData) WorkingCapitalMetrics {
	// DSO (Days Sales Outstanding)
	dsoValues := h.calculateDSO(data.Receivables, data.Revenue)
	dso := h.calculateEfficiencyMetrics(dsoValues, []float64{1, 1, 1, 1, 1}) // 日数なので分母は1
	
	// DIO (Days Inventory Outstanding)
	dioValues := h.calculateDIO(data.Inventory, data.COGS)
	dio := h.calculateEfficiencyMetrics(dioValues, []float64{1, 1, 1, 1, 1})
	
	// DPO (Days Payable Outstanding)
	dpoValues := h.calculateDPO(data.Payables, data.COGS)
	dpo := h.calculateEfficiencyMetrics(dpoValues, []float64{1, 1, 1, 1, 1})
	
	// Cash Cycle
	cashCycleValues := h.calculateCashCycle(dsoValues, dioValues, dpoValues)
	cashCycle := h.calculateEfficiencyMetrics(cashCycleValues, []float64{1, 1, 1, 1, 1})
	
	// Working Capital Intensity
	wcIntensityValues := h.calculateWCIntensity(data.Receivables, data.Inventory, data.Payables, data.Revenue)
	wcIntensity := h.calculateEfficiencyMetrics(wcIntensityValues, []float64{1, 1, 1, 1, 1})

	return WorkingCapitalMetrics{
		DSO:         dso,
		DIO:         dio,
		DPO:         dpo,
		CashCycle:   cashCycle,
		WCIntensity: wcIntensity,
	}
}

// Helper calculation methods

// calculateCAGR calculates Compound Annual Growth Rate
func (h *HistoricalAnalyzer) calculateCAGR(values []float64, years int) float64 {
	if len(values) < years+1 || values[0] == 0 {
		return 0.0
	}
	
	startValue := values[0]
	endValue := values[years]
	
	if startValue <= 0 || endValue <= 0 {
		return 0.0
	}
	
	return math.Pow(endValue/startValue, 1.0/float64(years)) - 1.0
}

// calculateAverage calculates the average of a slice
func (h *HistoricalAnalyzer) calculateAverage(values []float64) float64 {
	if len(values) == 0 {
		return 0.0
	}
	
	sum := 0.0
	for _, v := range values {
		sum += v
	}
	
	return sum / float64(len(values))
}

// calculateVolatility calculates volatility (standard deviation)
func (h *HistoricalAnalyzer) calculateVolatility(values []float64) float64 {
	if len(values) < 2 {
		return 0.0
	}
	
	mean := h.calculateAverage(values)
	sumSquares := 0.0
	
	for _, v := range values {
		diff := v - mean
		sumSquares += diff * diff
	}
	
	variance := sumSquares / float64(len(values)-1)
	return math.Sqrt(variance)
}

// determineTrendDirection determines the trend direction
func (h *HistoricalAnalyzer) determineTrendDirection(values []float64) string {
	if len(values) < 2 {
		return "stable"
	}
	
	positiveChanges := 0
	negativeChanges := 0
	
	for i := 1; i < len(values); i++ {
		if values[i] > values[i-1] {
			positiveChanges++
		} else if values[i] < values[i-1] {
			negativeChanges++
		}
	}
	
	if positiveChanges > negativeChanges {
		return "improving"
	} else if negativeChanges > positiveChanges {
		return "declining"
	}
	
	return "stable"
}

// calculateDSO calculates Days Sales Outstanding
func (h *HistoricalAnalyzer) calculateDSO(receivables, revenue []float64) []float64 {
	var dsoValues []float64
	
	for i := 0; i < len(receivables) && i < len(revenue); i++ {
		if revenue[i] != 0 {
			dso := (receivables[i] / revenue[i]) * 365
			dsoValues = append(dsoValues, dso)
		}
	}
	
	return dsoValues
}

// calculateDIO calculates Days Inventory Outstanding
func (h *HistoricalAnalyzer) calculateDIO(inventory, cogs []float64) []float64 {
	var dioValues []float64
	
	for i := 0; i < len(inventory) && i < len(cogs); i++ {
		if cogs[i] != 0 {
			dio := (inventory[i] / cogs[i]) * 365
			dioValues = append(dioValues, dio)
		}
	}
	
	return dioValues
}

// calculateDPO calculates Days Payable Outstanding
func (h *HistoricalAnalyzer) calculateDPO(payables, cogs []float64) []float64 {
	var dpoValues []float64
	
	for i := 0; i < len(payables) && i < len(cogs); i++ {
		if cogs[i] != 0 {
			dpo := (payables[i] / cogs[i]) * 365
			dpoValues = append(dpoValues, dpo)
		}
	}
	
	return dpoValues
}

// calculateCashCycle calculates Cash Conversion Cycle
func (h *HistoricalAnalyzer) calculateCashCycle(dso, dio, dpo []float64) []float64 {
	var cashCycleValues []float64
	
	minLen := len(dso)
	if len(dio) < minLen {
		minLen = len(dio)
	}
	if len(dpo) < minLen {
		minLen = len(dpo)
	}
	
	for i := 0; i < minLen; i++ {
		cashCycle := dso[i] + dio[i] - dpo[i]
		cashCycleValues = append(cashCycleValues, cashCycle)
	}
	
	return cashCycleValues
}

// calculateWCIntensity calculates Working Capital Intensity
func (h *HistoricalAnalyzer) calculateWCIntensity(receivables, inventory, payables, revenue []float64) []float64 {
	var wcIntensityValues []float64
	
	for i := 0; i < len(receivables) && i < len(inventory) && i < len(payables) && i < len(revenue); i++ {
		if revenue[i] != 0 {
			workingCapital := receivables[i] + inventory[i] - payables[i]
			intensity := workingCapital / revenue[i]
			wcIntensityValues = append(wcIntensityValues, intensity)
		}
	}
	
	return wcIntensityValues
}

// assessDataQuality assesses the quality of historical data
func (h *HistoricalAnalyzer) assessDataQuality(data *HistoricalData) QualityAssessment {
	totalDataPoints := len(data.Years) * 9 // 9 key metrics
	missingDataPoints := 0
	outlierCount := 0
	
	// データの完全性チェック
	dataArrays := [][]float64{
		data.Revenue, data.GrossProfit, data.EBITDA, data.NetIncome,
		data.TotalAssets, data.Receivables, data.Inventory, data.Payables, data.COGS,
	}
	
	for _, arr := range dataArrays {
		for _, val := range arr {
			if val == 0 || math.IsNaN(val) {
				missingDataPoints++
			}
		}
	}
	
	// 外れ値の検出
	for _, arr := range dataArrays {
		outlierCount += h.detectOutliers(arr)
	}
	
	dataCompleteness := float64(totalDataPoints-missingDataPoints) / float64(totalDataPoints)
	dataConsistency := 1.0 - (float64(outlierCount) / float64(totalDataPoints))
	reliabilityScore := (dataCompleteness + dataConsistency) / 2.0
	
	var recommendations []string
	if dataCompleteness < 0.9 {
		recommendations = append(recommendations, "データの欠損が多いため、追加のデータ収集を推奨")
	}
	if dataConsistency < 0.8 {
		recommendations = append(recommendations, "外れ値が多く検出されたため、データの検証を推奨")
	}
	if reliabilityScore < 0.7 {
		recommendations = append(recommendations, "データ品質が低いため、予測の信頼性に注意")
	}

	return QualityAssessment{
		DataCompleteness:   dataCompleteness,
		DataConsistency:    dataConsistency,
		OutlierCount:       outlierCount,
		ReliabilityScore:   reliabilityScore,
		RecommendedActions: recommendations,
	}
}

// detectOutliers detects outliers in a data series
func (h *HistoricalAnalyzer) detectOutliers(values []float64) int {
	if len(values) < 4 {
		return 0
	}
	
	sorted := make([]float64, len(values))
	copy(sorted, values)
	sort.Float64s(sorted)
	
	q1 := sorted[len(sorted)/4]
	q3 := sorted[3*len(sorted)/4]
	iqr := q3 - q1
	
	lowerBound := q1 - 1.5*iqr
	upperBound := q3 + 1.5*iqr
	
	outlierCount := 0
	for _, val := range values {
		if val < lowerBound || val > upperBound {
			outlierCount++
		}
	}
	
	return outlierCount
}

// generateForecastSuggestions generates forecast suggestions
func (h *HistoricalAnalyzer) generateForecastSuggestions(data *HistoricalData, growth GrowthAnalysis, margin MarginAnalysis) []ForecastSuggestion {
	var suggestions []ForecastSuggestion
	
	// 売上成長率の提案
	suggestions = append(suggestions, ForecastSuggestion{
		Category:       "Growth",
		Metric:         "Revenue Growth Rate",
		SuggestedValue: growth.RevenueGrowth.CAGR5Year,
		Confidence:     0.8,
		Rationale:      fmt.Sprintf("Based on 5-year CAGR of %.1f%%", growth.RevenueGrowth.CAGR5Year*100),
		RiskFactors:    []string{"Market conditions", "Competition", "Economic downturn"},
	})
	
	// 売上総利益率の提案
	suggestions = append(suggestions, ForecastSuggestion{
		Category:       "Margins",
		Metric:         "Gross Margin",
		SuggestedValue: margin.GrossMargin.Average,
		Confidence:     0.75,
		Rationale:      fmt.Sprintf("Based on historical average of %.1f%%", margin.GrossMargin.Average*100),
		RiskFactors:    []string{"Cost inflation", "Pricing pressure", "Product mix changes"},
	})
	
	// EBITDA利益率の提案
	suggestions = append(suggestions, ForecastSuggestion{
		Category:       "Margins",
		Metric:         "EBITDA Margin",
		SuggestedValue: margin.EBITDAMargin.Average,
		Confidence:     0.7,
		Rationale:      fmt.Sprintf("Based on historical average of %.1f%%", margin.EBITDAMargin.Average*100),
		RiskFactors:    []string{"Operating leverage", "Fixed cost inflation", "Efficiency improvements"},
	})
	
	return suggestions
}

// compareToBenchmarks compares metrics to industry benchmarks
func (h *HistoricalAnalyzer) compareToBenchmarks(data *HistoricalData, margin MarginAnalysis) map[string]interface{} {
	benchmarks := h.getIndustryBenchmarks(*h.params.IndustryType)
	
	return map[string]interface{}{
		"gross_margin_vs_benchmark": map[string]interface{}{
			"company_margin":    margin.GrossMargin.Average,
			"industry_benchmark": benchmarks["gross_margin"],
			"performance":       h.comparePerformance(margin.GrossMargin.Average, benchmarks["gross_margin"]),
		},
		"ebitda_margin_vs_benchmark": map[string]interface{}{
			"company_margin":     margin.EBITDAMargin.Average,
			"industry_benchmark": benchmarks["ebitda_margin"],
			"performance":        h.comparePerformance(margin.EBITDAMargin.Average, benchmarks["ebitda_margin"]),
		},
	}
}

// getIndustryBenchmarks returns industry-specific benchmarks
func (h *HistoricalAnalyzer) getIndustryBenchmarks(industryType string) map[string]float64 {
	benchmarks := map[string]map[string]float64{
		"technology": {
			"gross_margin":  0.70,
			"ebitda_margin": 0.25,
		},
		"manufacturing": {
			"gross_margin":  0.35,
			"ebitda_margin": 0.12,
		},
		"retail": {
			"gross_margin":  0.40,
			"ebitda_margin": 0.08,
		},
		"financial": {
			"gross_margin":  0.60,
			"ebitda_margin": 0.20,
		},
	}
	
	if industry, exists := benchmarks[industryType]; exists {
		return industry
	}
	
	// デフォルト値
	return benchmarks["manufacturing"]
}

// comparePerformance compares performance to benchmark
func (h *HistoricalAnalyzer) comparePerformance(actual, benchmark float64) string {
	if actual > benchmark*1.1 {
		return "Above Average"
	} else if actual < benchmark*0.9 {
		return "Below Average"
	}
	return "Average"
}

// calculateOverallConfidence calculates overall confidence
func (h *HistoricalAnalyzer) calculateOverallConfidence(suggestions []ForecastSuggestion) float64 {
	if len(suggestions) == 0 {
		return 0.0
	}
	
	totalConfidence := 0.0
	for _, suggestion := range suggestions {
		totalConfidence += suggestion.Confidence
	}
	
	return totalConfidence / float64(len(suggestions))
}

// outputResults outputs results to Excel
func (h *HistoricalAnalyzer) outputResults(result *ExcelHistoricalAnalysisResult) error {
	// 出力シートを作成
	err := h.file.CreateNewSheet(*h.params.OutputSheet)
	if err != nil {
		// シートが既に存在する場合は無視
	}
	outputSheet, err := h.file.FindSheet(*h.params.OutputSheet)
	if err != nil {
		return err
	}

	// 結果の出力
	row := 1
	outputSheet.SetValue(fmt.Sprintf("A%d", row), "Historical Analysis Results")
	row += 2
	
	// 分析サマリー
	outputSheet.SetValue(fmt.Sprintf("A%d", row), "Analysis Summary")
	row++
	for key, value := range result.AnalysisSummary {
		outputSheet.SetValue(fmt.Sprintf("A%d", row), key)
		outputSheet.SetValue(fmt.Sprintf("B%d", row), fmt.Sprintf("%v", value))
		row++
	}
	
	row += 2
	
	// 成長分析結果
	outputSheet.SetValue(fmt.Sprintf("A%d", row), "Growth Analysis")
	row++
	outputSheet.SetValue(fmt.Sprintf("A%d", row), "Revenue CAGR (5Y)")
	outputSheet.SetValue(fmt.Sprintf("B%d", row), fmt.Sprintf("%.2f%%", result.GrowthAnalysis.RevenueGrowth.CAGR5Year*100))
	row++
	outputSheet.SetValue(fmt.Sprintf("A%d", row), "Trend Direction")
	outputSheet.SetValue(fmt.Sprintf("B%d", row), result.GrowthAnalysis.OverallTrend)
	row++
	
	row += 2
	
	// 予測提案
	outputSheet.SetValue(fmt.Sprintf("A%d", row), "Forecast Suggestions")
	row++
	for _, suggestion := range result.ForecastSuggestions {
		outputSheet.SetValue(fmt.Sprintf("A%d", row), suggestion.Metric)
		outputSheet.SetValue(fmt.Sprintf("B%d", row), fmt.Sprintf("%.2f%%", suggestion.SuggestedValue*100))
		outputSheet.SetValue(fmt.Sprintf("C%d", row), fmt.Sprintf("%.0f%%", suggestion.Confidence*100))
		outputSheet.SetValue(fmt.Sprintf("D%d", row), suggestion.Rationale)
		row++
	}

	return nil
}

// getNumericValue gets a numeric value from a cell
func (h *HistoricalAnalyzer) getNumericValue(worksheet excel.Worksheet, cell string) float64 {
	value, err := worksheet.GetValue(cell)
	if err != nil {
		return 0.0
	}

	if floatVal, err := strconv.ParseFloat(strings.TrimSpace(value), 64); err == nil {
		return floatVal
	}

	return 0.0
}

// Additional helper methods for comprehensive analysis

// determineOverallTrend determines overall trend from multiple metrics
func (h *HistoricalAnalyzer) determineOverallTrend(metrics []GrowthMetrics) string {
	improvingCount := 0
	decliningCount := 0
	
	for _, metric := range metrics {
		switch metric.TrendDirection {
		case "improving":
			improvingCount++
		case "declining":
			decliningCount++
		}
	}
	
	if improvingCount > decliningCount {
		return "improving"
	} else if decliningCount > improvingCount {
		return "declining"
	}
	
	return "stable"
}

// calculateSeasonality calculates seasonality index
func (h *HistoricalAnalyzer) calculateSeasonality(values []float64) float64 {
	// 簡略化された季節性指数の計算
	if len(values) < 4 {
		return 0.0
	}
	
	// 実際の実装では、より複雑な季節性分析が必要
	return 0.1 // プレースホルダー
}

// determineMarginTrend determines margin trend from multiple margin metrics
func (h *HistoricalAnalyzer) determineMarginTrend(gross, ebitda, net MarginMetrics) string {
	trends := []string{gross.Trend, ebitda.Trend, net.Trend}
	
	improvingCount := 0
	decliningCount := 0
	
	for _, trend := range trends {
		switch trend {
		case "improving":
			improvingCount++
		case "declining":
			decliningCount++
		}
	}
	
	if improvingCount > decliningCount {
		return "improving"
	} else if decliningCount > improvingCount {
		return "declining"
	}
	
	return "stable"
}

// calculateMarginStability calculates margin stability
func (h *HistoricalAnalyzer) calculateMarginStability(gross, ebitda, net MarginMetrics) float64 {
	avgVolatility := (gross.Volatility + ebitda.Volatility + net.Volatility) / 3.0
	return 1.0 - avgVolatility // 高い安定性 = 低いボラティリティ
}

// calculateOverallEfficiency calculates overall efficiency
func (h *HistoricalAnalyzer) calculateOverallEfficiency(wc WorkingCapitalMetrics, assetTurnover EfficiencyMetrics) float64 {
	// 運転資本効率とアセットターンオーバーの加重平均
	wcScore := 1.0 / (1.0 + wc.CashCycle.Current/365.0) // 現金回転日数が短いほど効率的
	assetScore := assetTurnover.Current
	
	return (wcScore + assetScore) / 2.0
}

func AddExcelHistoricalAnalysisTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_historical_analysis",
		mcp.WithDescription("Perform comprehensive historical financial analysis with growth, margin, and efficiency metrics"),
		mcp.WithString("file_path",
			mcp.Required(),
			mcp.Description("Path to the Excel file containing historical data"),
		),
		mcp.WithString("data_sheet",
			mcp.Description("Name of the sheet containing historical data (default: Historical_Analysis)"),
		),
		mcp.WithNumber("start_year",
			mcp.Description("Start year for analysis (default: 2019)"),
		),
		mcp.WithNumber("end_year",
			mcp.Description("End year for analysis (default: 2023)"),
		),
		mcp.WithString("analysis_type",
			mcp.Description("Type of analysis (comprehensive, growth, margins, efficiency)"),
		),
		mcp.WithString("company_name",
			mcp.Description("Company name for analysis context"),
		),
		mcp.WithString("industry_type",
			mcp.Description("Industry type for benchmark comparison"),
		),
		mcp.WithString("output_sheet",
			mcp.Description("Name of output sheet (default: Analysis_Results)"),
		),
		mcp.WithBoolean("include_charts",
			mcp.Description("Include charts in analysis (default: true)"),
		),
		mcp.WithBoolean("include_forecasts",
			mcp.Description("Include forecast suggestions (default: true)"),
		),
	), func(arguments mcp.ToolCallArguments) *mcp.CallToolResult {
		var args ExcelHistoricalAnalysisParams
		if err := arguments.Unmarshal(&args); err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("Invalid arguments: %v", err))
		}

		result, err := ExcelHistoricalAnalysis(context.Background(), args)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("Failed to perform historical analysis: %v", err))
		}

		return mcp.NewToolResultText(fmt.Sprintf("Historical analysis completed: %s", result.Message))
	})
}