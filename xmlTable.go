package xlsx

type xlsxTable struct {
	Ref             string `xml:"ref,attr"`
	TotalsRowCount int    `xml:"totalsRowCount,attr,omitempty"`
	TableStyleInfo  xlsxTableStyleInfo `xml:"tableStyleInfo"`
}

type xlsxTableStyleInfo struct {
	Name string `xml:"name,attr"`
	ShowFirstColumn   int `xml:"showFirstColumn,attr"`
	ShowLastColumn    int `xml:"showLastColumn,attr"`
	ShowRowStripes    int `xml:"showRowStripes,attr"`
	ShowColumnStripes int `xml:"showColumnStripes,attr"`
}