# 三表統合財務モデル構築MCPサーバー開発計画書

## 🎯 **プロジェクト概要**

### 目的
投資銀行業界標準に準拠した三表統合財務モデル（損益計算書・貸借対照表・キャッシュフロー計算書）を自動生成できるMCPサーバーツールの開発

### 基盤技術
- **参考フレームワーク**: Paul Pignataro著『Financial Modeling and Valuation』メソドロジー
- **実装基準**: Otis Elevator Company ケーススタディによる実証済み手法
- **品質保証**: ウォール街投資銀行グレードの精度とフォーマット

---

## 🏗️ **MCPサーバーアーキテクチャ設計**

### **Tool 1: `financial_model_generator`**
三表統合モデルの完全自動生成エンジン

#### 入力パラメータ
```typescript
interface ModelGeneratorParams {
  company_name: string;
  historical_data: HistoricalFinancials;
  projection_years: number; // デフォルト: 5年
  model_type: 'standard' | 'acquisition' | 'lbo' | 'dcf_base';
  assumptions: ModelAssumptions;
  formatting_standard: 'pignataro' | 'otis' | 'custom';
}

interface HistoricalFinancials {
  income_statement: IncomeStatementData[];
  balance_sheet: BalanceSheetData[];
  cash_flow: CashFlowData[];
  years: number; // 最低3年、推奨5年
}

interface ModelAssumptions {
  revenue_drivers: RevenueDrivers;
  cost_structure: CostAssumptions;
  working_capital: WorkingCapitalDrivers;
  capital_policy: CapitalAllocationPolicy;
  market_environment: MarketAssumptions;
}
```

#### 出力仕様
```typescript
interface ModelOutput {
  excel_file_path: string;
  model_validation: ValidationResults;
  key_metrics: FinancialMetrics;
  scenario_analysis: ScenarioResults;
  documentation: ModelDocumentation;
}
```

### **Tool 2: `driver_based_revenue_calculator`**
Otis手法に基づく収益予測エンジン

#### 実装ロジック
```excel
// 市場規模成長率算定
Market_Growth_Rate = GDP_Growth × Industry_Beta + Demographic_Trends + Technology_Impact

// 市場シェア進展モデル
Market_Share_Progression = MIN(
  Current_Share + (Competitive_Advantage_Score × Investment_Level),
  Market_Share_Ceiling
)

// サービス事業収益算定
Service_Revenue = Service_Units × Revenue_Per_Unit × (1 + Service_Growth_Rate)

// 統合収益予測
Total_Revenue = (Market_Size × Market_Share) + Service_Revenue + Other_Revenue_Streams
```

### **Tool 3: `working_capital_optimizer`**
運転資本自動最適化ツール

#### 最適化アルゴリズム
```python
def optimize_working_capital(historical_ratios, industry_benchmarks, company_strategy):
    """
    DSO、DIO、DPO の最適バランスを算定
    キャッシュコンバージョンサイクルの最小化を目指す
    """
    target_dso = calculate_optimal_dso(
        customer_mix=company_strategy.customer_profile,
        collection_efficiency=historical_ratios.collection_trends,
        industry_standard=industry_benchmarks.dso_median
    )
    
    target_dio = optimize_inventory_days(
        business_model=company_strategy.inventory_strategy,
        supply_chain_efficiency=historical_ratios.inventory_turns,
        seasonal_patterns=historical_ratios.seasonality
    )
    
    target_dpo = maximize_payable_days(
        supplier_relationships=company_strategy.supplier_power,
        early_payment_benefits=historical_ratios.discount_capture,
        cash_flow_constraints=company_strategy.liquidity_requirements
    )
    
    return WorkingCapitalTargets(target_dso, target_dio, target_dpo)
```

### **Tool 4: `circular_reference_resolver`**
循環参照自動解決エンジン

#### 解決手法
```typescript
interface CircularReferenceResolver {
  debt_interest_calculation: 'beginning_balance' | 'average_balance' | 'iterative';
  cash_management: 'automatic_sweep' | 'target_balance' | 'policy_driven';
  convergence_criteria: {
    max_iterations: number; // デフォルト: 100
    tolerance: number; // デフォルト: 0.001
  };
}

// 実装例：債務利息計算
function resolve_interest_expense(debt_schedule: DebtData[], rates: InterestRates[]): number[] {
  return debt_schedule.map((debt, index) => {
    const beginning_balance = index === 0 ? debt.opening_balance : debt_schedule[index-1].ending_balance;
    return beginning_balance * rates[index].interest_rate;
  });
}
```

### **Tool 5: `financial_ratio_analyzer`**
財務比率自動分析・ベンチマーキングツール

#### 分析対象指標
```typescript
interface FinancialRatios {
  profitability: {
    gross_margin: number;
    ebitda_margin: number;
    net_margin: number;
    roe: number;
    roa: number;
    roic: number;
  };
  
  liquidity: {
    current_ratio: number;
    quick_ratio: number;
    cash_ratio: number;
    operating_cash_flow_ratio: number;
  };
  
  leverage: {
    debt_to_equity: number;
    debt_to_assets: number;
    interest_coverage: number;
    ebitda_coverage: number;
  };
  
  efficiency: {
    asset_turnover: number;
    inventory_turnover: number;
    dso: number;
    dpo: number;
    cash_conversion_cycle: number;
  };
}
```

### **Tool 6: `scenario_generator`**
シナリオ分析・感応度テスト自動実行

#### シナリオ設定
```typescript
interface ScenarioDefinition {
  base_case: ModelAssumptions;
  optimistic: {
    revenue_growth_premium: number; // +20%
    margin_expansion: number; // +100bps
    market_share_gains: number; // +50bps
  };
  pessimistic: {
    revenue_decline: number; // -15%
    margin_compression: number; // -150bps
    competitive_pressure: number; // -100bps market share
  };
  stress_test: {
    recession_impact: RecessionParameters;
    interest_rate_shock: RateShockParameters;
    competitive_disruption: DisruptionParameters;
  };
}
```

---

## 📊 **Excel統合仕様**

### **Sheet構造標準化**
```typescript
interface ExcelWorkbookStructure {
  sheets: {
    executive_summary: ExecutiveSummarySheet;
    assumptions: AssumptionsSheet;
    historical_analysis: HistoricalDataSheet;
    income_statement: IncomeStatementSheet;
    balance_sheet: BalanceSheetSheet;
    cash_flow: CashFlowSheet;
    working_capital: WorkingCapitalSheet;
    debt_schedule: DebtScheduleSheet;
    depreciation: DepreciationSheet;
    scenarios: ScenarioComparisonSheet;
    model_checks: ValidationSheet;
    charts_dashboard: VisualizationSheet;
  };
  
  formatting: {
    color_scheme: IndustryStandardColors;
    font_standards: FontSpecification;
    number_formats: NumberFormattingRules;
    conditional_formatting: ConditionalFormattingRules;
  };
}
```

### **数式標準化ライブラリ**
```excel
// 収益成長率計算
Revenue_Growth = (Current_Year_Revenue / Prior_Year_Revenue) - 1

// 運転資本変動計算  
WC_Change = -(∆AR + ∆Inventory - ∆AP - ∆Accrued_Liabilities)

// ROIC計算
ROIC = NOPAT / Average_Invested_Capital
NOPAT = EBIT × (1 - Tax_Rate)
Invested_Capital = Working_Capital + Net_Fixed_Assets + Goodwill

// フリーキャッシュフロー計算
FCF = EBIT × (1 - Tax_Rate) + Depreciation - CapEx - ∆Working_Capital

// 企業価値計算
Enterprise_Value = Equity_Value + Net_Debt + Preferred_Stock + NCI
```

---

## 🔧 **技術実装計画**

### **Phase 1: コア機能開発 (4週間)**

#### Week 1-2: 基盤インフラ構築
```python
# MCPサーバー基本構造
class ThreeStatementModelServer:
    def __init__(self):
        self.excel_engine = ExcelEngine()
        self.formula_library = FormulaLibrary()
        self.validation_engine = ValidationEngine()
        self.formatting_engine = FormattingEngine()
    
    async def generate_model(self, params: ModelGeneratorParams) -> ModelOutput:
        # 1. 履歴データ検証・標準化
        validated_data = await self.validate_historical_data(params.historical_data)
        
        # 2. ドライバー分析・前提条件設定
        drivers = await self.analyze_business_drivers(validated_data)
        
        # 3. 三表予測モデル構築
        model = await self.build_integrated_model(drivers, params.assumptions)
        
        # 4. バランシング・検証
        validated_model = await self.validate_model_integrity(model)
        
        # 5. フォーマット適用・出力
        formatted_output = await self.apply_professional_formatting(validated_model)
        
        return formatted_output
```

#### Week 3-4: 数式エンジン開発
```python
class FormulaEngine:
    def __init__(self):
        self.formula_templates = self.load_pignataro_templates()
        self.otis_patterns = self.load_otis_patterns()
    
    def generate_income_statement_formulas(self, drivers: BusinessDrivers) -> Dict[str, str]:
        """損益計算書数式自動生成"""
        formulas = {}
        
        # 収益予測数式
        formulas['product_revenue'] = f"={drivers.market_size_cell}*{drivers.market_share_cell}"
        formulas['service_revenue'] = f"={drivers.service_units_cell}*{drivers.revenue_per_unit_cell}"
        formulas['total_revenue'] = f"=SUM({formulas['product_revenue']},{formulas['service_revenue']})"
        
        # コスト予測数式
        formulas['product_cogs'] = f"=-{formulas['product_revenue']}*{drivers.product_cogs_rate_cell}"
        formulas['service_cogs'] = f"=-{formulas['service_revenue']}*{drivers.service_cogs_rate_cell}"
        
        return formulas
    
    def generate_balance_sheet_formulas(self, income_refs: Dict, drivers: BusinessDrivers) -> Dict[str, str]:
        """貸借対照表数式自動生成"""
        formulas = {}
        
        # 運転資本数式
        formulas['accounts_receivable'] = f"={income_refs['total_revenue']}/365*{drivers.dso_cell}"
        formulas['inventory'] = f"=({income_refs['product_cogs']}+{income_refs['service_cogs']})/365*{drivers.inventory_days_cell}"
        formulas['accounts_payable'] = f"=({income_refs['product_cogs']}+{income_refs['service_cogs']})/365*{drivers.dpo_cell}"
        
        return formulas
```

### **Phase 2: 高度機能実装 (4週間)**

#### Week 5-6: シナリオ分析機能
```python
class ScenarioAnalysisEngine:
    def generate_scenario_models(self, base_model: FinancialModel, scenarios: ScenarioDefinition) -> Dict[str, FinancialModel]:
        scenario_models = {}
        
        for scenario_name, scenario_params in scenarios.items():
            # 前提条件調整
            adjusted_assumptions = self.adjust_assumptions(base_model.assumptions, scenario_params)
            
            # モデル再計算
            scenario_model = self.recalculate_model(base_model, adjusted_assumptions)
            
            # 結果比較テーブル生成
            comparison_table = self.generate_comparison_table(base_model, scenario_model)
            
            scenario_models[scenario_name] = {
                'model': scenario_model,
                'comparison': comparison_table,
                'key_variances': self.calculate_key_variances(base_model, scenario_model)
            }
        
        return scenario_models
```

#### Week 7-8: 品質保証・バリデーション機能
```python
class ModelValidationEngine:
    def validate_model_integrity(self, model: FinancialModel) -> ValidationResults:
        checks = []
        
        # 1. 貸借対照表バランスチェック
        balance_check = self.check_balance_sheet_balance(model.balance_sheet)
        checks.append(balance_check)
        
        # 2. キャッシュフロー整合性チェック
        cf_check = self.check_cash_flow_integrity(model.cash_flow, model.balance_sheet)
        checks.append(cf_check)
        
        # 3. 三表連携チェック
        integration_check = self.check_statement_integration(model)
        checks.append(integration_check)
        
        # 4. 合理性チェック
        reasonableness_check = self.check_assumption_reasonableness(model.assumptions)
        checks.append(reasonableness_check)
        
        return ValidationResults(checks)
    
    def check_balance_sheet_balance(self, balance_sheet: BalanceSheetData) -> ValidationCheck:
        """貸借対照表バランス検証"""
        for period in balance_sheet.periods:
            total_assets = sum(period.assets.values())
            total_liab_equity = sum(period.liabilities.values()) + sum(period.equity.values())
            
            if abs(total_assets - total_liab_equity) > 0.01:
                return ValidationCheck(
                    name="Balance Sheet Balance",
                    status="FAIL",
                    message=f"Period {period.date}: Assets {total_assets} != Liab+Equity {total_liab_equity}",
                    severity="CRITICAL"
                )
        
        return ValidationCheck(name="Balance Sheet Balance", status="PASS", severity="INFO")
```

### **Phase 3: ユーザーインターフェース (2週間)**

#### Week 9-10: Claude統合・ユーザビリティ向上
```python
class ClaudeIntegrationLayer:
    def __init__(self, model_server: ThreeStatementModelServer):
        self.model_server = model_server
        self.conversation_context = ConversationContext()
    
    async def handle_model_request(self, user_request: str, attached_files: List[str]) -> str:
        """Claude経由のモデル構築リクエスト処理"""
        
        # 1. ユーザー意図解析
        intent = await self.parse_user_intent(user_request)
        
        # 2. 添付ファイル解析
        historical_data = await self.extract_historical_data(attached_files)
        
        # 3. 前提条件推定・確認
        assumptions = await self.infer_assumptions(intent, historical_data)
        confirmation = await self.confirm_assumptions_with_user(assumptions)
        
        # 4. モデル生成
        model_output = await self.model_server.generate_model(
            ModelGeneratorParams(
                historical_data=historical_data,
                assumptions=confirmation.final_assumptions,
                model_type=intent.model_type
            )
        )
        
        # 5. 結果説明・推奨事項生成
        explanation = await self.generate_model_explanation(model_output)
        
        return explanation
    
    async def generate_model_explanation(self, model_output: ModelOutput) -> str:
        """モデル結果の分かりやすい説明生成"""
        explanation = f"""
        ## 📊 三表統合財務モデル構築完了
        
        ### 🎯 主要財務指標
        - **売上成長率**: {model_output.key_metrics.revenue_cagr:.1%}
        - **EBITDAマージン**: {model_output.key_metrics.avg_ebitda_margin:.1%}
        - **ROE**: {model_output.key_metrics.avg_roe:.1%}
        - **フリーキャッシュフロー**: ${model_output.key_metrics.total_fcf:,.0f}M
        
        ### ✅ モデル品質検証
        {self.format_validation_results(model_output.model_validation)}
        
        ### 📈 シナリオ分析結果
        {self.format_scenario_analysis(model_output.scenario_analysis)}
        
        ### 📁 出力ファイル
        完成したExcelモデル: `{model_output.excel_file_path}`
        
        ### 💡 推奨事項
        {self.generate_recommendations(model_output)}
        """
        
        return explanation
```

---

## 🎛️ **設定・カスタマイゼーション**

### **設定ファイル構造**
```yaml
# mcp_config.yaml
three_statement_model_server:
  default_settings:
    projection_years: 5
    model_type: "standard"
    formatting_standard: "pignataro"
    validation_strictness: "high"
    
  excel_settings:
    color_scheme: "professional_blue"
    font_family: "Calibri"
    decimal_places: 0
    currency_symbol: "$"
    
  calculation_settings:
    circular_reference_resolution: "beginning_balance"
    convergence_tolerance: 0.001
    max_iterations: 100
    
  industry_benchmarks:
    default_industry: "Industrial"
    benchmark_sources: ["Bloomberg", "CapitalIQ", "FactSet"]
    
  validation_rules:
    balance_sheet_tolerance: 0.01
    growth_rate_limits: [-50%, 100%]
    margin_reasonableness: [0%, 50%]
    
  output_preferences:
    include_charts: true
    generate_executive_summary: true
    create_scenario_analysis: true
    export_formats: ["xlsx", "pdf"]
```

### **業界別テンプレート**
```python
INDUSTRY_TEMPLATES = {
    'industrial': {
        'working_capital_patterns': {
            'dso_range': [30, 60],
            'inventory_days_range': [45, 90],
            'dpo_range': [30, 45]
        },
        'margin_benchmarks': {
            'gross_margin': [25%, 40%],
            'ebitda_margin': [10%, 20%],
            'net_margin': [5%, 15%]
        },
        'capex_intensity': [1%, 3%],
        'debt_capacity': [2.0, 4.0]  # EBITDA multiples
    },
    
    'technology': {
        'working_capital_patterns': {
            'dso_range': [45, 75],
            'inventory_days_range': [30, 60],
            'dpo_range': [45, 60]
        },
        'margin_benchmarks': {
            'gross_margin': [60%, 80%],
            'ebitda_margin': [20%, 40%],
            'net_margin': [15%, 30%]
        },
        'capex_intensity': [2%, 5%],
        'debt_capacity': [1.0, 3.0]
    }
}
```

---

## 🚀 **開発工程表**

### **マイルストーン**

| フェーズ | 期間 | 主要成果物 | 検証項目 |
|---------|------|------------|----------|
| **Phase 1** | Week 1-4 | コアMCPサーバー、基本三表生成機能 | Otisケースの完全再現 |
| **Phase 2** | Week 5-8 | シナリオ分析、高度検証機能 | 複数企業での精度検証 |
| **Phase 3** | Week 9-10 | Claude統合、UI/UX最適化 | エンドツーエンド動作確認 |
| **Testing** | Week 11-12 | 総合テスト、パフォーマンス最適化 | 本番環境リリース準備 |

### **成功指標**

#### 技術的KPI
- **精度**: 手動モデルとの差異 < 0.1%
- **パフォーマンス**: 5年予測モデル生成時間 < 30秒
- **信頼性**: バリデーションテスト合格率 > 99%

#### ビジネス的KPI  
- **時間短縮**: 従来の手動作業時間を80%削減
- **品質向上**: 計算エラー発生率を95%削減
- **使いやすさ**: 金融専門知識レベルに関係なく利用可能

---

## 🔍 **リスク管理・品質保証**

### **主要リスクと対策**

#### 技術的リスク
1. **数式複雑性**: 段階的実装、十分なテストカバレッジ
2. **Excel互換性**: 複数バージョンでの動作確認
3. **パフォーマンス**: プロファイリング、最適化の継続実施

#### ビジネスリスク
1. **業界標準乖離**: 金融専門家レビュー、業界ベンチマーク比較
2. **規制要件**: コンプライアンス専門家との連携
3. **ユーザー受容性**: 段階的ロールアウト、フィードバック反映

### **品質保証プロセス**
```python
class QualityAssuranceFramework:
    def run_comprehensive_qa(self, model_output: ModelOutput) -> QAResults:
        qa_tests = [
            self.numerical_accuracy_test(model_output),
            self.formula_integrity_test(model_output),
            self.professional_formatting_test(model_output),
            self.business_logic_test(model_output),
            self.performance_test(model_output),
            self.user_acceptance_test(model_output)
        ]
        
        return QAResults(qa_tests)
```

---

## 📈 **将来展開計画**

### **Phase 4: 高度分析機能 (将来)**
- DCF評価モデル統合
- Monte Carlo シミュレーション
- 機械学習による予測精度向上
- リアルタイム市場データ連携

### **Phase 5: エコシステム拡張**
- 他社分析ツール（Bloomberg、FactSet）との連携
- クラウドベース共同編集機能
- モバイルアプリケーション開発
- API エコシステム構築

---

## 💡 **実装開始における推奨アプローチ**

### **最初のMVP（Minimum Viable Product）**
1. Otisケーススタディの完全再現機能
2. 基本的な三表統合ロジック
3. 標準的な業界フォーマットサポート
4. Claudeとの基本的な対話インターフェース

### **段階的機能追加**
1. **Week 1-2**: コア数式エンジン
2. **Week 3-4**: Excel出力・フォーマット
3. **Week 5-6**: バリデーション・品質チェック
4. **Week 7-8**: シナリオ分析機能
5. **Week 9-10**: Claude統合・ユーザビリティ

この開発計画により、投資銀行グレードの三表統合財務モデルを自動生成できる包括的なMCPサーバーシステムを構築し、財務分析業務の効率化と精度向上を実現いたします。