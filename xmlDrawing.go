package xlsx

type xlsxWsDr struct {
	TwoCellAnchors []xlsxTwoCellAnchor `xml:"twoCellAnchor"`
}

type xlsxTwoCellAnchor struct {
	From xlsxPos  `xml:"from"`
	To   xlsxPos  `xml:"to"`
	Pic  *xlsxPic `xml:"pic"`
}

type xlsxPos struct {
	Col    int `xml:"col"`
	ColOff int `xml:"colOff"`
	Row    int `xml:"row"`
	RowOff int `xml:"rowOff"`
}

type xlsxPic struct {
	NvPicPr  xlsxNvPicPr  `xml:"nvPicPr"`
	BlipFill xlsxBlipFill `xml:"blipFill"`
	SpPr     xlsxSpPr     `xml:"spPr"`
}

type xlsxNvPicPr struct {
	CNvPr xlsxCNvPr `xml:"cNvPr"`
}

type xlsxCNvPr struct {
	Id    int    `xml:"id,attr"`
	Name  string `xml:"name,attr"`
	Descr string `xml:"descr,attr"`
}

type xlsxBlipFill struct {
	Blip xlsxBlip `xml:"blip"`
}

type xlsxBlip struct {
	Embed string `xml:"embed,attr"`
}

type xlsxSpPr struct {
	Xfrm xlsxXfrm `xml:"xfrm"`
}

type xlsxXfrm struct {
	Off xlsxOff `xml:"off"`
	Ext xlsxExt `xml:"ext"`
}

type xlsxOff struct {
	X int `xml:"x,attr"`
	Y int `xml:"y,attr"`
}

type xlsxExt struct {
	CX int `xml:"cx,attr"`
	CY int `xml:"cy,attr"`
}
