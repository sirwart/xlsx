package xlsx

type xlsxTable struct {
	Ref             string `xml:"ref,attr"`
	TotalsRowCount  int    `xml:"totalsRowCount,attr,omitempty"`
	TableStyleInfo  xlsxTableStyleInfo `xml:"tableStyleInfo"`
}

type xlsxTableStyleInfo struct {
	Name string `xml:"name,attr"`
	ShowFirstColumn   int `xml:"showFirstColumn,attr"`
	ShowLastColumn    int `xml:"showLastColumn,attr"`
	ShowRowStripes    int `xml:"showRowStripes,attr"`
	ShowColumnStripes int `xml:"showColumnStripes,attr"`
}

type xlsxPivotTableDefinition struct {
	Location xlsxLocation `xml:"location"`
	RowItems xlsxRowItems `xml:"rowItems"`
	PivotTableStyleInfo xlsxPivotTableStyleInfo `xml:"pivotTableStyleInfo"`
}

type xlsxLocation struct {
	Ref string `xml:"ref,attr"`
}

type xlsxRowItems struct {
	Items []xlsxRowItem `xml:"i"`
}

type xlsxRowItem struct {
	T string `xml:"t,attr,omitempty"`
}

type xlsxPivotTableStyleInfo struct {
	Name           string `xml:"name,attr"`
	ShowRowStripes int    `xml:"showRowStripes,attr"`
	ShowColStripes int    `xml:"showColStripes,attr"`
}
