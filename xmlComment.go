package xlsx

type xlsxComments struct {
	CommentList xlsxCommentList `xml:"commentList"`
}

type xlsxCommentList struct {
	Comments []xlsxComment `xml:"comment"`
}

type xlsxComment struct {
	Ref  string   `xml:"ref,attr"`
	Text xlsxText `xml:"text"`
}

type xlsxText struct {
	R []xlsxRichText `xml:"r"`
}

type xlsxRichText struct {
	T string `xml:"t"`
}
