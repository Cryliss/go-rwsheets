package rwsheets

import (
	"errors"
	"time"

	sheets "google.golang.org/api/sheets/v4"
)

// GetSheetData: Retrieve the spreadsheet data.
func GetSheetData(ssid, readRange string, srv *sheets.Service) ([]*sheets.RowData, error) {
	var rows []*sheets.RowData
	var ranges []string
	ranges = append(ranges, readRange)

	noData := errors.New("no sheet data found")

	// Get the spreadsheet data
	ss, err := srv.Spreadsheets.Get(ssid).Ranges(ranges...).IncludeGridData(true).Do()
	if err != nil {
		return rows, err
	}

	// Make sure we actually got at least one sheet of data
	if len(ss.Sheets) == 0 {
		return rows, noData
	}

	// Extract the row data from the sheet
	sheet := ss.Sheets[0]

	// Make sure we actually have data
	if len(sheet.Data) == 0 {
		return rows, noData
	}

	grid := sheet.Data[0]
	rows = grid.RowData

	return rows, nil
}

// UpdateSheetData: Update the spreadsheet with new values.
func UpdateSheetData(ssid string, endColumnIndex, gid, startColumnIndex, startRowIndex int64, newVals []*sheets.RowData, srv *sheets.Service) error {
	var batchUpdate sheets.BatchUpdateSpreadsheetRequest
	batchUpdate.IncludeSpreadsheetInResponse = false

	gridRange := sheets.GridRange{
		EndColumnIndex:   endColumnIndex,
		SheetId:          gid,
		StartColumnIndex: startColumnIndex,
		StartRowIndex:    startRowIndex,
		EndRowIndex:      startRowIndex + int64(len(newVals)),
	}

	updateCells := sheets.UpdateCellsRequest{
		Fields: "*",
		Range:  &gridRange,
		Rows:   newVals,
	}

	request := sheets.Request{
		UpdateCells: &updateCells,
	}
	batchUpdate.Requests = append(batchUpdate.Requests, &request)
	batchUpdate.MarshalJSON()

	if _, err := srv.Spreadsheets.BatchUpdate(ssid, &batchUpdate).Do(); err != nil {
		return err
	}

	return nil
}

// RemoveRow: For removing a specific row in a Sheet.
func RemoveRow(rows []*sheets.RowData, rmvIdx int) []*sheets.RowData {
	if len(rows) == rmvIdx {
		// Let's assume we wanted to remove the last row ..
		return rows[:rmvIdx-2]
	}

	if rmvIdx+1 > len(rows) {
		return rows
	}

	return append(rows[:rmvIdx], rows[rmvIdx+1:]...)
}

// BoolValue: For updating UserEnteredValue with a boolean value.
func BoolValue(value bool) *sheets.ExtendedValue {
	return &sheets.ExtendedValue{
		BoolValue: &value,
	}
}

// FormulaValue: For updating UserEnteredValue with a formula value.
func FormulaValue(value string) *sheets.ExtendedValue {
	return &sheets.ExtendedValue{
		FormulaValue: &value,
	}
}

// TextValue: For updating UserEnteredValue with a text value.
func TextValue(value string) *sheets.ExtendedValue {
	return &sheets.ExtendedValue{
		StringValue: &value,
	}
}

// NumberValue: For updating UserEnteredValue with a number value.
// Should be used in combination with SerialDate for date values.
func NumberValue(value float64) *sheets.ExtendedValue {
	return &sheets.ExtendedValue{
		NumberValue: &value,
	}
}

// SerialDate: Returns the Google Sheets serial number for the date.
func SerialDate(value, format string) (float64, error) {
	newDate, err := time.Parse(format, value)
	if err != nil {
		return float64(0.0), err
	}

	startDate, err := time.Parse("1/2/2006", "12/30/1899")
	if err != nil {
		return float64(0.0), err
	}

	days := newDate.Sub(startDate).Hours() / 24
	return float64(days), nil
}
