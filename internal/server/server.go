package server

import (
	"runtime"

	"github.com/mark3labs/mcp-go/server"
	"github.com/negokaz/excel-mcp-server/internal/tools"
)

type ExcelServer struct {
	server *server.MCPServer
}

func New(version string) *ExcelServer {
	s := &ExcelServer{}
	s.server = server.NewMCPServer(
		"excel-mcp-server",
		version,
	)
	tools.AddExcelCreateFileTool(s.server)
	// tools.AddExcelDeleteFileTool(s.server)   // TODO: Excel interfaceにGetSheetNames/DeleteSheet追加が必要
	// tools.AddExcelDeleteSheetTool(s.server)  // TODO: Excel interfaceにGetSheetNames/DeleteSheet追加が必要
	tools.AddExcelDescribeSheetsTool(s.server)
	tools.AddExcelReadSheetTool(s.server)
	if runtime.GOOS == "windows" {
		tools.AddExcelScreenCaptureTool(s.server)
	}
	tools.AddExcelWriteToSheetTool(s.server)
	// tools.AddExcelFormatCellsTool(s.server)        // TODO: zog API更新が必要
	// tools.AddExcelDataValidationTool(s.server)     // TODO: zog API更新が必要
	// tools.AddExcelGoalSeekTool(s.server)           // TODO: zog API更新が必要
	// tools.AddExcelDataTableTool(s.server)          // TODO: zog API更新が必要
	// tools.AddExcelArrayLimiterTool(s.server)       // TODO: zog API更新が必要
	tools.AddExcelCreateTableTool(s.server)
	tools.AddExcelCopySheetTool(s.server)
	tools.AddExcelFormatRangeTool(s.server)
	// フェーズ1: 財務モデリング基盤機能 (TODO: zog API更新が必要)
	// tools.AddExcelFinancialTemplateTool(s.server)
	// tools.AddExcelCircularReferenceTool(s.server)
	// tools.AddExcelModelValidationTool(s.server)
	// tools.AddExcelFinancialFormattingTool(s.server)
	// tools.AddExcelHistoricalAnalysisTool(s.server)
	return s
}

func (s *ExcelServer) Start() error {
	return server.ServeStdio(s.server)
}
