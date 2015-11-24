package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"io/ioutil"
	"path"
	"strconv"
	"strings"
)

// XLSXReaderError is the standard error type for otherwise undefined
// errors in the XSLX reading process.
type XLSXReaderError struct {
	Err string
}

// String() returns a string value from an XLSXReaderError struct in
// order that it might comply with the os.Error interface.
func (e *XLSXReaderError) Error() string {
	return e.Err
}

// getRangeFromString is an internal helper function that converts
// XLSX internal range syntax to a pair of integers.  For example,
// the range string "1:3" yield the upper and lower intergers 1 and 3.
func getRangeFromString(rangeString string) (lower int, upper int, error error) {
	var parts []string
	parts = strings.SplitN(rangeString, ":", 2)
	if parts[0] == "" {
		error = errors.New(fmt.Sprintf("Invalid range '%s'\n", rangeString))
	}
	if parts[1] == "" {
		error = errors.New(fmt.Sprintf("Invalid range '%s'\n", rangeString))
	}
	lower, error = strconv.Atoi(parts[0])
	if error != nil {
		error = errors.New(fmt.Sprintf("Invalid range (not integer in lower bound) %s\n", rangeString))
	}
	upper, error = strconv.Atoi(parts[1])
	if error != nil {
		error = errors.New(fmt.Sprintf("Invalid range (not integer in upper bound) %s\n", rangeString))
	}
	return lower, upper, error
}

// LettersToNumeric is used to convert a character based column
// reference to a zero based numeric column identifier.
func LettersToNumeric(letters string) int {
	sum, mul, n := 0, 1, 0
	for i := len(letters) - 1; i >= 0; i, mul, n = i-1, mul*26, 1 {
		c := letters[i]
		switch {
		case 'A' <= c && c <= 'Z':
			n += int(c - 'A')
		case 'a' <= c && c <= 'z':
			n += int(c - 'a')
		}
		sum += n * mul
	}
	return sum
}

// Get the largestDenominator that is a multiple of a basedDenominator
// and fits at least once into a given numerator.
func getLargestDenominator(numerator, multiple, baseDenominator, power int) (int, int) {
	if numerator/multiple == 0 {
		return 1, power
	}
	next, nextPower := getLargestDenominator(
		numerator, multiple*baseDenominator, baseDenominator, power+1)
	if next > multiple {
		return next, nextPower
	}
	return multiple, power
}

// Convers a list of numbers representing a column into a alphabetic
// representation, as used in the spreadsheet.
func formatColumnName(colId []int) string {
	lastPart := len(colId) - 1

	result := ""
	for n, part := range colId {
		if n == lastPart {
			// The least significant number is in the
			// range 0-25, all other numbers are 1-26,
			// hence we use a differente offset for the
			// last part.
			result += string(part + 65)
		} else {
			// Don't output leading 0s, as there is no
			// representation of 0 in this format.
			if part > 0 {
				result += string(part + 64)
			}
		}
	}
	return result
}

func smooshBase26Slice(b26 []int) []int {
	// Smoosh values together, eliminating 0s from all but the
	// least significant part.
	lastButOnePart := len(b26) - 2
	for i := lastButOnePart; i > 0; i-- {
		part := b26[i]
		if part == 0 {
			greaterPart := b26[i-1]
			if greaterPart > 0 {
				b26[i-1] = greaterPart - 1
				b26[i] = 26
			}
		}
	}
	return b26
}

func intToBase26(x int) (parts []int) {
	// Excel column codes are pure evil - in essence they're just
	// base26, but they don't represent the number 0.
	b26Denominator, _ := getLargestDenominator(x, 1, 26, 0)

	// This loop terminates because integer division of 1 / 26
	// returns 0.
	for d := b26Denominator; d > 0; d = d / 26 {
		value := x / d
		remainder := x % d
		parts = append(parts, value)
		x = remainder
	}
	return parts
}

// numericToLetters is used to convert a zero based, numeric column
// indentifier into a character code.
func numericToLetters(colRef int) string {
	parts := intToBase26(colRef)
	return formatColumnName(smooshBase26Slice(parts))
}

// letterOnlyMapF is used in conjunction with strings.Map to return
// only the characters A-Z and a-z in a string
func letterOnlyMapF(rune rune) rune {
	switch {
	case 'A' <= rune && rune <= 'Z':
		return rune
	case 'a' <= rune && rune <= 'z':
		return rune - 32
	}
	return -1
}

// intOnlyMapF is used in conjunction with strings.Map to return only
// the numeric portions of a string.
func intOnlyMapF(rune rune) rune {
	if rune >= 48 && rune < 58 {
		return rune
	}
	return -1
}

// getCoordsFromCellIDString returns the zero based cartesian
// coordinates from a cell name in Excel format, e.g. the cellIDString
// "A1" returns 0, 0 and the "B3" return 1, 2.
func getCoordsFromCellIDString(cellIDString string) (x, y int, error error) {
	var letterPart string = strings.Map(letterOnlyMapF, cellIDString)
	y, error = strconv.Atoi(strings.Map(intOnlyMapF, cellIDString))
	if error != nil {
		return x, y, error
	}
	y -= 1 // Zero based
	x = LettersToNumeric(letterPart)
	return x, y, error
}

// getCellIDStringFromCoords returns the Excel format cell name that
// represents a pair of zero based cartesian coordinates.
func getCellIDStringFromCoords(x, y int) string {
	letterPart := numericToLetters(x)
	numericPart := y + 1
	return fmt.Sprintf("%s%d", letterPart, numericPart)
}

// getMaxMinFromDimensionRef return the zero based cartesian maximum
// and minimum coordinates from the dimension reference embedded in a
// XLSX worksheet.  For example, the dimension reference "A1:B2"
// returns "0,0", "1,1".
func getMaxMinFromDimensionRef(ref string) (minx, miny, maxx, maxy int, err error) {
	var parts []string
	parts = strings.Split(ref, ":")
	minx, miny, err = getCoordsFromCellIDString(parts[0])
	if err != nil {
		return -1, -1, -1, -1, err
	}
	if len(parts) == 1 {
		maxx, maxy = minx, miny
		return
	}
	maxx, maxy, err = getCoordsFromCellIDString(parts[1])
	if err != nil {
		return -1, -1, -1, -1, err
	}
	return
}

// calculateMaxMinFromWorkSheet works out the dimensions of a spreadsheet
// that doesn't have a DimensionRef set.  The only case currently
// known where this is true is with XLSX exported from Google Docs.
func calculateMaxMinFromWorksheet(worksheet *xlsxWorksheet) (minx, miny, maxx, maxy int, err error) {
	// Note, this method could be very slow for large spreadsheets.
	var x, y int
	var maxVal int
	maxVal = int(^uint(0) >> 1)
	minx = maxVal
	miny = maxVal
	maxy = 0
	maxx = 0
	for _, row := range worksheet.SheetData.Row {
		for _, cell := range row.C {
			x, y, err = getCoordsFromCellIDString(cell.R)
			if err != nil {
				return -1, -1, -1, -1, err
			}
			if x < minx {
				minx = x
			}
			if x > maxx {
				maxx = x
			}
			if y < miny {
				miny = y
			}
			if y > maxy {
				maxy = y
			}
		}
	}
	if minx == maxVal {
		minx = 0
	}
	if miny == maxVal {
		miny = 0
	}
	return
}

// makeRowFromSpan will, when given a span expressed as a string,
// return an empty Row large enough to encompass that span and
// populate it with empty cells.  All rows start from cell 1 -
// regardless of the lower bound of the span.
func makeRowFromSpan(spans string) *Row {
	var error error
	var upper int
	var row *Row
	var cell *Cell

	row = new(Row)
	_, upper, error = getRangeFromString(spans)
	if error != nil {
		panic(error)
	}
	error = nil
	row.Cells = make([]*Cell, upper)
	for i := 0; i < upper; i++ {
		cell = new(Cell)
		cell.Value = ""
		row.Cells[i] = cell
	}
	return row
}

// makeRowFromRaw returns the Row representation of the xlsxRow.
func makeRowFromRaw(rawrow xlsxRow) *Row {
	var upper int
	var row *Row
	var cell *Cell

	row = new(Row)
	upper = -1

	for _, rawcell := range rawrow.C {
		if rawcell.R != "" {
			x, _, error := getCoordsFromCellIDString(rawcell.R)
			if error != nil {
				panic(fmt.Sprintf("Invalid Cell Coord, %s\n", rawcell.R))
			}
			if x > upper {
				upper = x
			}
			continue
		}
		upper++
	}
	upper++

	row.Cells = make([]*Cell, upper)
	for i := 0; i < upper; i++ {
		cell = new(Cell)
		cell.Value = ""
		row.Cells[i] = cell
	}
	return row
}

func makeEmptyRow() *Row {
	row := new(Row)
	row.Cells = make([]*Cell, 0)
	return row
}

type sharedFormula struct {
	x, y    int
	formula string
}

func formulaForCell(rawcell xlsxC, sharedFormulas map[int]sharedFormula) string {
	var res string

	f := rawcell.F
	if f.T == "shared" {
		x, y, err := getCoordsFromCellIDString(rawcell.R)
		if err != nil {
			res = f.Content
		} else {
			if f.Ref != "" {
				res = f.Content
				sharedFormulas[f.Si] = sharedFormula{x, y, res}
			} else {
				sharedFormula := sharedFormulas[f.Si]
				dx := x - sharedFormula.x
				dy := y - sharedFormula.y
				orig := []byte(sharedFormula.formula)
				var start, end int
				for end = 0; end < len(orig); end++ {
					c := orig[end]
					if c >= 'A' && c <= 'Z' {
						res += string(orig[start:end])
						start = end
						end++
						foundNum := false
						for ; end < len(orig); end++ {
							idc := orig[end]
							if idc >= '0' && idc <= '9' {
								foundNum = true
							} else if idc >= 'A' && idc <= 'Z' {
								if foundNum {
									break
								}
							} else {
								break
							}
						}
						if foundNum {
							fx, fy, _ := getCoordsFromCellIDString(string(orig[start:end]))
							fx += dx
							fy += dy
							res += getCellIDStringFromCoords(fx, fy)
							start = end
						}
					}
				}
				if start < len(orig) {
					res += string(orig[start:end])
				}
			}
		}
	} else {
		res = f.Content
	}
	return strings.Trim(res, " \t\n\r")
}

// fillCellData attempts to extract a valid value, usable in
// CSV form from the raw cell value.  Note - this is not actually
// general enough - we should support retaining tabs and newlines.
func fillCellData(rawcell xlsxC, reftable *RefTable, sharedFormulas map[int]sharedFormula, cell *Cell) {
	var data string = rawcell.V
	if len(data) > 0 {
		vval := strings.Trim(data, " \t\n\r")
		switch rawcell.T {
		case "s": // Shared String
			ref, error := strconv.Atoi(vval)
			if error != nil {
				panic(error)
			}
			cell.Value = reftable.ResolveSharedString(ref)
			cell.cellType = CellTypeString
		case "b": // Boolean
			cell.Value = vval
			cell.cellType = CellTypeBool
		case "e": // Error
			cell.Value = vval
			if rawcell.F != nil {
				cell.formula = formulaForCell(rawcell, sharedFormulas)
			}
			cell.cellType = CellTypeError
		default:
			if rawcell.F == nil {
				// Numeric
				cell.Value = vval
				cell.cellType = CellTypeNumeric
			} else {
				// Formula
				cell.Value = vval
				cell.formula = formulaForCell(rawcell, sharedFormulas)
				cell.cellType = CellTypeFormula
			}
		}
	}
}

// readRowsFromSheet is an internal helper function that extracts the
// rows from a XSLXWorksheet, populates them with Cells and resolves
// the value references from the reference table and stores them in
// the rows and columns.
func readRowsFromSheet(Worksheet *xlsxWorksheet, file *File, rels map[string]xlsxWorkbookRelation) ([]*Row, []*Col, int, int) {
	var rows []*Row
	var cols []*Col
	var row *Row
	var minCol, maxCol, minRow, maxRow, colCount, rowCount int
	var reftable *RefTable
	var err error
	var insertRowIndex, insertColIndex int
	sharedFormulas := map[int]sharedFormula{}

	if len(Worksheet.SheetData.Row) == 0 {
		return nil, nil, 0, 0
	}
	reftable = file.referenceTable
	if len(Worksheet.Dimension.Ref) > 0 {
		minCol, minRow, maxCol, maxRow, err = getMaxMinFromDimensionRef(Worksheet.Dimension.Ref)
	} else {
		minCol, minRow, maxCol, maxRow, err = calculateMaxMinFromWorksheet(Worksheet)
	}
	if err != nil {
		panic(err.Error())
	}
	rowCount = maxRow + 1
	colCount = maxCol + 1
	rows = make([]*Row, rowCount)
	cols = make([]*Col, colCount)
	insertRowIndex = minRow
	for i := range cols {
		cols[i] = &Col{
			Hidden: false,
		}
	}

	hyperlinks := map[string]string{}
	if Worksheet.Hyperlinks != nil {
		for _, hyperlink := range Worksheet.Hyperlinks.Links {
			if hyperlink.Location != "" {
				hyperlinks[hyperlink.Ref] = hyperlink.Location
			} else if hyperlink.Id != "" {
				rel, ok := rels[hyperlink.Id]
				if ok && rel.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" {
					hyperlinks[hyperlink.Ref] = rel.Target
				}
			}
		}
	}

	getStyle := func(styleIndex int) *Style {
		if file.styles != nil {
			return file.styles.getStyle(styleIndex)
		} else {
			return nil
		}
	}

	// Columns can apply to a range, for convenience we expand the
	// ranges out into individual column definitions.
	for _, rawcol := range Worksheet.Cols.Col {
		// Note, below, that sometimes column definitions can
		// exist outside the defined dimensions of the
		// spreadsheet - we deliberately exclude these
		// columns.
		for i := rawcol.Min; i <= rawcol.Max && i <= colCount; i++ {
			cols[i-1] = &Col{
				Min:    rawcol.Min,
				Max:    rawcol.Max,
				Hidden: rawcol.Hidden,
				Width:  rawcol.Width,
				Style:  getStyle(rawcol.Style)}
		}
	}

	// insert leading empty rows that is in front of minRow
	for rowIndex := 0; rowIndex < minRow; rowIndex++ {
		rows[rowIndex] = makeEmptyRow()
	}

	for rowIndex := 0; rowIndex < len(Worksheet.SheetData.Row); rowIndex++ {
		rawrow := Worksheet.SheetData.Row[rowIndex]
		// Some spreadsheets will omit blank rows from the
		// stored data
		for rawrow.R > (insertRowIndex + 1) {
			// Put an empty Row into the array
			rows[insertRowIndex] = makeEmptyRow()
			insertRowIndex++
		}
		// range is not empty and only one range exist
		if len(rawrow.Spans) != 0 && strings.Count(rawrow.Spans, ":") == 1 {
			row = makeRowFromSpan(rawrow.Spans)
		} else {
			row = makeRowFromRaw(rawrow)
		}

		row.Hidden = rawrow.Hidden
		row.Height = rawrow.Ht
		row.Style = getStyle(rawrow.S)

		insertColIndex = minCol
		for _, rawcell := range rawrow.C {
			x, _, _ := getCoordsFromCellIDString(rawcell.R)

			// Some spreadsheets will omit blank cells
			// from the data.
			for x > insertColIndex {
				// Put an empty Cell into the array
				row.Cells[insertColIndex] = new(Cell)
				insertColIndex++
			}
			cellX := insertColIndex
			cell := row.Cells[cellX]
			fillCellData(rawcell, reftable, sharedFormulas, cell)
			if file.styles != nil {
				cell.style = file.styles.getStyle(rawcell.S)
				cell.numFmt = file.styles.getNumberFormat(rawcell.S)
			}
			cell.date1904 = file.Date1904
			cell.Hidden = rawrow.Hidden || (len(cols) > cellX && cell.Hidden)
			cell.Href = hyperlinks[rawcell.R]
			insertColIndex++
		}
		if len(rows) > insertRowIndex {
			rows[insertRowIndex] = row
		}
		insertRowIndex++
	}
	return rows, cols, colCount, rowCount
}

type indexedSheet struct {
	Index int
	Sheet *Sheet
	Error error
}

func readSheetViews(xSheetViews xlsxSheetViews) []SheetView {
	if xSheetViews.SheetView == nil || len(xSheetViews.SheetView) == 0 {
		return nil
	}
	sheetViews := []SheetView{}
	for _, xSheetView := range xSheetViews.SheetView {
		sheetView := SheetView{ShowGridLines: true}
		if xSheetView.ShowGridLines != nil {
			sheetView.ShowGridLines = *xSheetView.ShowGridLines
		}
		if xSheetView.Pane != nil {
			xlsxPane := xSheetView.Pane
			pane := &Pane{}
			pane.XSplit = xlsxPane.XSplit
			pane.YSplit = xlsxPane.YSplit
			pane.TopLeftCell = xlsxPane.TopLeftCell
			pane.ActivePane = xlsxPane.ActivePane
			pane.State = xlsxPane.State
			sheetView.Pane = pane
		}
		sheetViews = append(sheetViews, sheetView)
	}
	return sheetViews
}

func readMergeCells(xCells []xlsxMergeCell) (cells []MergeCell) {
	for _, xCell := range xCells {
		refs := strings.Split(xCell.Ref, ":")
		if len(refs) != 2 {
			continue
		}

		cell := MergeCell{refs[0], refs[1]}
		cells = append(cells, cell)
	}
	return
}

// readSheetFromFile is the logic of converting a xlsxSheet struct
// into a Sheet struct.  This work can be done in parallel and so
// readWorksheetRelationsFromZipFile will spawn an instance of this function per
// sheet and get the results back on the provided channel.
func readSheetFromFile(sc chan *indexedSheet, index int, rsheet xlsxSheet, fi *File, sheetXMLMap map[string]string, worksheetRels, relableFiles map[string]*zip.File) {
	result := &indexedSheet{Index: index, Sheet: nil, Error: nil}
	worksheet, err := getWorksheetFromSheet(rsheet, fi.worksheets, sheetXMLMap)
	if err != nil {
		result.Error = err
		sc <- result
		return
	}
	sheet := new(Sheet)
	sheet.SheetId = rsheet.SheetId
	sheet.File = fi

	var rels map[string]xlsxWorkbookRelation
	if worksheetRelFile := worksheetFileForSheet(rsheet, worksheetRels, sheetXMLMap); worksheetRelFile != nil {
		rels, err = readWorksheetRelationsFromZipFile(worksheetRelFile)
	}

	sheet.Rows, sheet.Cols, sheet.MaxCol, sheet.MaxRow = readRowsFromSheet(worksheet, fi, rels)
	sheet.Hidden = rsheet.State == sheetStateHidden || rsheet.State == sheetStateVeryHidden
	sheet.SheetViews = readSheetViews(worksheet.SheetViews)

	sheet.SheetFormat.DefaultColWidth = worksheet.SheetFormatPr.DefaultColWidth
	sheet.SheetFormat.DefaultRowHeight = worksheet.SheetFormatPr.DefaultRowHeight

	if mergeCells := worksheet.MergeCells; mergeCells != nil {
		sheet.MergeCells = readMergeCells(mergeCells.Cells)
	}

	sheet.HasDrawing = worksheet.Drawing != nil

	for _, rel := range rels {
		filename := rel.Target[3:len(rel.Target)]
		if file := relableFiles[filename]; file != nil {
			switch rel.Type {
			case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments":
				comments, err := readCommentsFromCommentFile(file)
				if err == nil {
					sheet.Comments = comments
				}
			case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table":
				table, err := readTableFromFile(file)
				if err == nil {
					sheet.Tables = append(sheet.Tables, *table)
				}
			case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable":
				pivotTable, err := readPivotTableFromFile(file)
				if err == nil {
					sheet.PivotTables = append(sheet.PivotTables, *pivotTable)
				}
			case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing":
				components := strings.Split(filename, "/")
				lastComponent := components[len(components)-1]
				lastComponent += ".rels"
				components[len(components)-1] = "_rels"
				components = append(components, lastComponent)
				relsFilename := strings.Join(components, "/")
				relsFile := relableFiles[relsFilename]
				if relsFile == nil {
					continue
				}
				rels, err := readWorksheetRelationsFromZipFile(relsFile)
				if err != nil {
					continue
				}

				drawings, err := readDrawingsFromFile(file, rels, relableFiles)
				if err == nil {
					sheet.Drawings = drawings
				}
			}
		}
	}

	result.Sheet = sheet
	sc <- result
}

// readSheetsFromZipFile is an internal helper function that loops
// over the Worksheets defined in the XSLXWorkbook and loads them into
// Sheet objects stored in the Sheets slice of a xlsx.File struct.
func readSheetsFromZipFile(f *zip.File, file *File, sheetXMLMap map[string]string, worksheetRels, relableFiles map[string]*zip.File) (map[string]*Sheet, []*Sheet, error) {
	var workbook *xlsxWorkbook
	var err error
	var rc io.ReadCloser
	var decoder *xml.Decoder
	var sheetCount int
	workbook = new(xlsxWorkbook)
	rc, err = f.Open()
	if err != nil {
		return nil, nil, err
	}
	decoder = xml.NewDecoder(rc)
	err = decoder.Decode(workbook)
	if err != nil {
		return nil, nil, err
	}
	file.Date1904 = workbook.WorkbookPr.Date1904

	// Only try and read sheets that have corresponding files.
	// Notably this excludes chartsheets don't right now
	var workbookSheets []xlsxSheet
	for _, sheet := range workbook.Sheets.Sheet {
		if f := worksheetFileForSheet(sheet, file.worksheets, sheetXMLMap); f != nil {
			workbookSheets = append(workbookSheets, sheet)
		}
	}
	sheetCount = len(workbookSheets)
	sheetsByName := make(map[string]*Sheet, sheetCount)
	sheets := make([]*Sheet, sheetCount)
	sheetChan := make(chan *indexedSheet, sheetCount)
	defer close(sheetChan)

	go func() {
		defer func() {
			if e := recover(); e != nil {
				err = fmt.Errorf("%v", e)
				result := &indexedSheet{Index: -1, Sheet: nil, Error: err}
				sheetChan <- result
			}
		}()
		err = nil
		for i, rawsheet := range workbookSheets {
			readSheetFromFile(sheetChan, i, rawsheet, file, sheetXMLMap, worksheetRels, relableFiles)
		}
	}()

	for j := 0; j < sheetCount; j++ {
		sheet := <-sheetChan
		if sheet.Error != nil {
			return nil, nil, sheet.Error
		}
		sheetName := workbookSheets[sheet.Index].Name
		sheetsByName[sheetName] = sheet.Sheet
		sheet.Sheet.Name = sheetName
		sheets[sheet.Index] = sheet.Sheet
	}
	return sheetsByName, sheets, nil
}

// readSharedStringsFromZipFile() is an internal helper function to
// extract a reference table from the sharedStrings.xml file within
// the XLSX zip file.
func readSharedStringsFromZipFile(f *zip.File) (*RefTable, error) {
	var sst *xlsxSST
	var error error
	var rc io.ReadCloser
	var decoder *xml.Decoder
	var reftable *RefTable

	// In a file with no strings it's possible that
	// sharedStrings.xml doesn't exist.  In this case the value
	// passed as f will be nil.
	if f == nil {
		return nil, nil
	}
	rc, error = f.Open()
	if error != nil {
		return nil, error
	}
	sst = new(xlsxSST)
	decoder = xml.NewDecoder(rc)
	error = decoder.Decode(sst)
	if error != nil {
		return nil, error
	}
	reftable = MakeSharedStringRefTable(sst)
	return reftable, nil
}

// readStylesFromZipFile() is an internal helper function to
// extract a style table from the style.xml file within
// the XLSX zip file.
func readStylesFromZipFile(f *zip.File, theme *theme) (*xlsxStyleSheet, error) {
	var style *xlsxStyleSheet
	var error error
	var rc io.ReadCloser
	var decoder *xml.Decoder
	rc, error = f.Open()
	if error != nil {
		return nil, error
	}
	style = newXlsxStyleSheet(theme)
	decoder = xml.NewDecoder(rc)
	error = decoder.Decode(style)
	if error != nil {
		return nil, error
	}
	buildNumFmtRefTable(style)
	return style, nil
}

func buildNumFmtRefTable(style *xlsxStyleSheet) {
	for _, numFmt := range style.NumFmts.NumFmt {
		// We do this for the side effect of populating the NumFmtRefTable.
		style.addNumFmt(numFmt)
	}
}

func readThemeFromZipFile(f *zip.File) (*theme, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}

	var themeXml xlsxTheme
	err = xml.NewDecoder(rc).Decode(&themeXml)
	if err != nil {
		return nil, err
	}

	return newTheme(themeXml), nil
}

type WorkBookRels map[string]string

func (w *WorkBookRels) MakeXLSXWorkbookRels() xlsxWorkbookRels {
	relCount := len(*w)
	xWorkbookRels := xlsxWorkbookRels{}
	xWorkbookRels.Relationships = make([]xlsxWorkbookRelation, relCount+3)
	for k, v := range *w {
		index, err := strconv.Atoi(k[3:])
		if err != nil {
			panic(err.Error())
		}
		xWorkbookRels.Relationships[index-1] = xlsxWorkbookRelation{
			Id:     k,
			Target: v,
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"}
	}

	relCount++
	sheetId := fmt.Sprintf("rId%d", relCount)
	xWorkbookRels.Relationships[relCount-1] = xlsxWorkbookRelation{
		Id:     sheetId,
		Target: "sharedStrings.xml",
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"}

	relCount++
	sheetId = fmt.Sprintf("rId%d", relCount)
	xWorkbookRels.Relationships[relCount-1] = xlsxWorkbookRelation{
		Id:     sheetId,
		Target: "theme/theme1.xml",
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"}

	relCount++
	sheetId = fmt.Sprintf("rId%d", relCount)
	xWorkbookRels.Relationships[relCount-1] = xlsxWorkbookRelation{
		Id:     sheetId,
		Target: "styles.xml",
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"}

	return xWorkbookRels
}

// readWorkbookRelationsFromZipFile is an internal helper function to
// extract a map of relationship ID strings to the name of the
// worksheet.xml file they refer to.  The resulting map can be used to
// reliably derefence the worksheets in the XLSX file.
func readWorkbookRelationsFromZipFile(workbookRels *zip.File) (WorkBookRels, error) {
	var sheetXMLMap WorkBookRels
	var wbRelationships *xlsxWorkbookRels
	var rc io.ReadCloser
	var decoder *xml.Decoder
	var err error

	rc, err = workbookRels.Open()
	if err != nil {
		return nil, err
	}
	decoder = xml.NewDecoder(rc)
	wbRelationships = new(xlsxWorkbookRels)
	err = decoder.Decode(wbRelationships)
	if err != nil {
		return nil, err
	}
	sheetXMLMap = make(WorkBookRels)
	for _, rel := range wbRelationships.Relationships {
		if strings.HasSuffix(rel.Target, ".xml") && rel.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" {
			_, filename := path.Split(rel.Target)
			sheetXMLMap[rel.Id] = strings.Replace(filename, ".xml", "", 1)
		}
	}
	return sheetXMLMap, nil
}

func readWorksheetRelationsFromZipFile(worksheetRels *zip.File) (map[string]xlsxWorkbookRelation, error) {
	rc, err := worksheetRels.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	rels := xlsxWorkbookRels{}
	err = xml.NewDecoder(rc).Decode(&rels)
	if err != nil {
		return nil, err
	}
	res := map[string]xlsxWorkbookRelation{}
	for _, relation := range rels.Relationships {
		res[relation.Id] = relation
	}
	return res, nil
}

func readCommentsFromCommentFile(commentFile *zip.File) ([]Comment, error) {
	rc, err := commentFile.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	comments := xlsxComments{}
	err = xml.NewDecoder(rc).Decode(&comments)
	if err != nil {
		return nil, err
	}
	res := []Comment{}
	for _, comment := range comments.CommentList.Comments {
		var text string
		for _, r := range comment.Text.R {
			text += r.T
		}
		res = append(res, Comment{comment.Ref, text})
	}
	return res, nil
}

func readTableFromFile(file *zip.File) (*Table, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	xTable := xlsxTable{}
	err = xml.NewDecoder(rc).Decode(&xTable)
	if err != nil {
		return nil, err
	}

	styleInfo := xTable.TableStyleInfo
	tableStyleInfo := TableStyleInfo{styleInfo.Name, styleInfo.ShowFirstColumn != 0, styleInfo.ShowLastColumn != 0,
									 styleInfo.ShowRowStripes != 0,	styleInfo.ShowColumnStripes != 0}

	refs := strings.Split(xTable.Ref, ":")
	if len(refs) != 2 {
		return nil, errors.New("Invalid table ref: "+xTable.Ref)
	}

	table := Table{refs[0], refs[1], xTable.TotalsRowCount, tableStyleInfo}
	return &table, nil
}

func readPivotTableFromFile(file *zip.File) (*PivotTable, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	xPivotTable := xlsxPivotTableDefinition{}
	err = xml.NewDecoder(rc).Decode(&xPivotTable)
	if err != nil {
		return nil, err
	}

	styleInfo := xPivotTable.PivotTableStyleInfo
	pivotTableStyleInfo := PivotTableStyleInfo{styleInfo.Name, styleInfo.ShowRowStripes != 0,	styleInfo.ShowColStripes != 0}

	refs := strings.Split(xPivotTable.Location.Ref, ":")
	if len(refs) != 2 {
		return nil, errors.New("Invalid table ref: "+xPivotTable.Location.Ref)
	}

	var rowItems []RowItem

	for _, rowItem := range xPivotTable.RowItems.Items {
		rowItems = append(rowItems, RowItem{rowItem.T})
	}

	pivotTable := PivotTable{refs[0], refs[1], rowItems, pivotTableStyleInfo}
	return &pivotTable, nil
}

func readDrawingsFromFile(file *zip.File, rels map[string]xlsxWorkbookRelation, relableFiles map[string]*zip.File) ([]Drawing, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	wsDr := xlsxWsDr{}
	err = xml.NewDecoder(rc).Decode(&wsDr)
	if err != nil {
		return nil, err
	}

	var drawings []Drawing

	for _, anchor := range wsDr.TwoCellAnchors {
		if anchor.Pic != nil {
			embed := anchor.Pic.BlipFill.Blip.Embed
			if embed == "" {
				continue
			}
			picRel := rels[embed]
			imageFilename := picRel.Target[3:len(picRel.Target)]
			imageFile := relableFiles[imageFilename]
			imageReader, err := imageFile.Open()
			if err != nil {
				continue
			}
			defer imageReader.Close()
			image, err := ioutil.ReadAll(imageReader)
			if err != nil {
				continue
			}
			pic := Pic{Image: image}

			xfrm := anchor.Pic.SpPr.Xfrm

			drawing := Drawing{xfrm.Off.X, xfrm.Off.Y, xfrm.Ext.CX, xfrm.Ext.CY, pic}
			drawings = append(drawings, drawing)
		}
	}

	return drawings, nil
}

// ReadZip() takes a pointer to a zip.ReadCloser and returns a
// xlsx.File struct populated with its contents.  In most cases
// ReadZip is not used directly, but is called internally by OpenFile.
func ReadZip(f *zip.ReadCloser) (*File, error) {
	defer f.Close()
	return ReadZipReader(&f.Reader)
}

// ReadZipReader() can be used to read an XLSX in memory without
// touching the filesystem.
func ReadZipReader(r *zip.Reader) (*File, error) {
	var err error
	var file *File
	var reftable *RefTable
	var sharedStrings *zip.File
	var sheetXMLMap map[string]string
	var sheetsByName map[string]*Sheet
	var sheets []*Sheet
	var style *xlsxStyleSheet
	var styles *zip.File
	var themeFile *zip.File
	var v *zip.File
	var workbook *zip.File
	var workbookRels *zip.File
	var worksheets map[string]*zip.File

	file = NewFile()
	// file.numFmtRefTable = make(map[int]xlsxNumFmt, 1)
	worksheets = make(map[string]*zip.File, len(r.File))
	worksheetRels := map[string]*zip.File{}
	relableFiles := map[string]*zip.File{}
	for _, v = range r.File {
		switch v.Name {
		case "xl/sharedStrings.xml":
			sharedStrings = v
		case "xl/workbook.xml":
			workbook = v
		case "xl/_rels/workbook.xml.rels":
			workbookRels = v
		case "xl/styles.xml":
			styles = v
		case "xl/theme/theme1.xml":
			themeFile = v
		default:
			if strings.HasPrefix(v.Name, "xl/comments") || strings.HasPrefix(v.Name, "xl/tables") ||
			   strings.HasPrefix(v.Name, "xl/pivotTables") || strings.HasPrefix(v.Name, "xl/drawings") ||
			   strings.HasPrefix(v.Name, "xl/media") {
				relableFiles[v.Name[3:len(v.Name)]] = v
			} else if len(v.Name) > 29 && v.Name[0:20] == "xl/worksheets/_rels/" {
				worksheetRels[v.Name[20:len(v.Name)-9]] = v
			} else if len(v.Name) > 14 {
				if v.Name[0:13] == "xl/worksheets" {
					worksheets[v.Name[14:len(v.Name)-4]] = v
				}
			}
		}
	}
	sheetXMLMap, err = readWorkbookRelationsFromZipFile(workbookRels)
	if err != nil {
		return nil, err
	}
	file.worksheets = worksheets
	reftable, err = readSharedStringsFromZipFile(sharedStrings)
	if err != nil {
		return nil, err
	}
	file.referenceTable = reftable
	if themeFile != nil {
		theme, err := readThemeFromZipFile(themeFile)
		if err != nil {
			return nil, err
		}

		file.theme = theme
	}
	if styles != nil {
		style, err = readStylesFromZipFile(styles, file.theme)
		if err != nil {
			return nil, err
		}

		file.styles = style
	}
	sheetsByName, sheets, err = readSheetsFromZipFile(workbook, file, sheetXMLMap, worksheetRels, relableFiles)
	if err != nil {
		return nil, err
	}
	if sheets == nil {
		readerErr := new(XLSXReaderError)
		readerErr.Err = "No sheets found in XLSX File"
		return nil, readerErr
	}
	file.Sheet = sheetsByName
	file.Sheets = sheets
	return file, nil
}
