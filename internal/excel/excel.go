package excel

import (
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

type Excel interface {
	// GetBackendName returns the backend used to manipulate the Excel file.
	GetBackendName() string
	// GetSheets returns a list of all worksheets in the Excel file.
	GetSheets() ([]Worksheet, error)
	// FindSheet finds a sheet by its name and returns a Worksheet.
	FindSheet(sheetName string) (Worksheet, error)
	// CreateNewSheet creates a new sheet with the specified name.
	CreateNewSheet(sheetName string) error
	// CopySheet copies a sheet from one to another.
	CopySheet(srcSheetName, destSheetName string) error
	// Save saves the Excel file.
	Save() error
	// FormatCells applies formatting to a range of cells
	// セルの範囲にフォーマットを適用するのです！ ✨(ﾉ◕ヮ◕)ﾉ*:･ﾟ✧
	FormatCells(sheetName, rangeStr string, style map[string]interface{}) error
}

type Worksheet interface {
	// Release releases the worksheet resources.
	Release()
	// Name returns the name of the worksheet.
	Name() (string, error)
	// GetTable returns a tables in this worksheet.
	GetTables() ([]Table, error)
	// GetPivotTable returns a pivot tables in this worksheet.
	GetPivotTables() ([]PivotTable, error)
	// SetValue sets a value in the specified cell.
	SetValue(cell string, value any) error
	// SetFormula sets a formula in the specified cell.
	SetFormula(cell string, formula string) error
	// GetValue gets the value from the specified cell.
	GetValue(cell string) (string, error)
	// GetFormula gets the formula from the specified cell.
	GetFormula(cell string) (string, error)
	// GetDimention gets the dimension of the worksheet.
	GetDimention() (string, error)
	// GetPagingStrategy returns the paging strategy for the worksheet.
	// The pageSize parameter is used to determine the max size of each page.
	GetPagingStrategy(pageSize int) (PagingStrategy, error)
	// CapturePicture returns base64 encoded image data of the specified range.
	CapturePicture(captureRange string) (string, error)
	// AddTable adds a table to this worksheet.
	AddTable(tableRange, tableName string) error
}

type Table struct {
	Name  string
	Range string
}

type PivotTable struct {
	Name  string
	Range string
}

// OpenFile opens an Excel file and returns an Excel interface.
// It first tries to open the file using OLE automation, and if that fails,
// it tries to using the excelize library.
// If the file doesn't exist, it creates a new Excel file.
func OpenFile(absoluteFilePath string) (Excel, func(), error) {
	// Check if file exists first
	fileExists := true
	if _, err := os.Stat(absoluteFilePath); os.IsNotExist(err) {
		fileExists = false
	}
	
	if fileExists {
		// Try OLE for existing files
		ole, releaseFn, err := NewExcelOle(absoluteFilePath)
		if err == nil {
			return ole, releaseFn, nil
		}
	} else {
		// Try OLE for new files (Windows only)
		ole, releaseFn, err := NewExcelOleWithNewFile(absoluteFilePath)
		if err == nil {
			return ole, releaseFn, nil
		}
	}
	
	// If OLE fails or not available, use Excelize
	return OpenOrCreateFile(absoluteFilePath)
}

// OpenOrCreateFile opens an existing Excel file or creates a new one if it doesn't exist.
func OpenOrCreateFile(absoluteFilePath string) (Excel, func(), error) {
	// Check if file exists
	if _, err := os.Stat(absoluteFilePath); os.IsNotExist(err) {
		// File doesn't exist, create a new one
		return CreateNewFile(absoluteFilePath)
	}
	
	// File exists, try to open it
	workbook, err := excelize.OpenFile(absoluteFilePath)
	if err != nil {
		return nil, func() {}, err
	}
	excel := NewExcelizeExcel(workbook)
	return excel, func() {
		workbook.Close()
	}, nil
}

// CreateNewFile creates a new Excel file at the specified path.
func CreateNewFile(absoluteFilePath string) (Excel, func(), error) {
	// Ensure the directory exists
	dir := filepath.Dir(absoluteFilePath)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return nil, func() {}, err
	}
	
	// Create new workbook
	workbook := excelize.NewFile()
	
	// Set the file path for saving
	workbook.Path = absoluteFilePath
	
	excel := NewExcelizeExcel(workbook)
	return excel, func() {
		workbook.Close()
	}, nil
}
