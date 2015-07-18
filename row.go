package xlsx

type Row struct {
	Cells  []*Cell
	Hidden bool
	Height float32
	Sheet  *Sheet
}

func (r *Row) AddCell() *Cell {
	cell := NewCell(r)
	r.Cells = append(r.Cells, cell)
	r.Sheet.maybeAddCol(len(r.Cells))
	return cell
}
