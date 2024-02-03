package main

import (
	"context"
	"encoding/json"
	"log"
	"os"
	"strconv"

	rwsheets "github.com/cryliss/go-rwsheets"
	"github.com/joho/godotenv"
	sheets "google.golang.org/api/sheets/v4"
)

func init() {
	// Load the environment file
	if err := godotenv.Load(); err != nil {
		log.Fatalf("Error loading env file: %v", err)
	}
}

// Updates a spreadsheet with data found in a JSON file and applies formatting to the cells.
func main() {
	credentialFile := os.Getenv("CREDENTIALS")
	tokenFile := os.Getenv("TOKEN")
	scopes := "https://www.googleapis.com/auth/spreadsheets"

	srv, err := rwsheets.NewSheetsService(context.Background(), credentialFile, tokenFile, scopes)
	if err != nil {
		log.Fatalf("failed to create a new sheets service - %s", err.Error())
	}

	ssid := os.Getenv("SSID")
	gidStr := os.Getenv("GID")
	gid, err := strconv.ParseInt(gidStr, 10, 64)
	if err != nil {
		log.Fatalf("failed to parse sheet GID into int64 from environment vairable - %s", err.Error())
	}

	data, err := getSampleData(os.Getenv("SAMPLE_DATA"))
	if err != nil {
		log.Fatalf("failed to read sample data from file - %s", err.Error())
	}

	styler := rwsheets.NewStyler().FontBold(true).FontFamily("Verdana").FontSize(int64(12)).HorizontalAlignment("CENTER").VerticalAlignment("MIDDLE")

	headerBorders := rwsheets.BorderConf{
		Bottom: true,
		Left:   true,
		Right:  true,
		Style:  "SOLID_MEDIUM",
		Top:    true,
	}
	newRows := styler.CreateHeaderRow(data.Headers, &headerBorders)

	cellBorders := rwsheets.BorderConf{
		Bottom: true,
		Left:   true,
		Right:  true,
		Style:  "SOLID",
		Top:    false,
	}
	dateLayout := "2006-01-02"
	styler.FontSize(int64(10)).FontBold(false)

	for i, invoice := range data.Invoices {
		if i > 0 {
			cellBorders.Top = true
		}
		var cells []*sheets.CellData
		cells = append(cells, styler.HorizontalAlignment("LEFT").TextCell(invoice.Customer, &cellBorders))
		cells = append(cells, styler.HorizontalAlignment("CENTER").TextCell(invoice.Invoice, &cellBorders))
		cells = append(cells, styler.HorizontalAlignment("RIGHT").AccountingCell(invoice.Amount, &cellBorders))
		cells = append(cells, styler.HorizontalAlignment("RIGHT").DatePattern("M/d/yyyy").DateCell(invoice.Date, dateLayout, &cellBorders))
		cells = append(cells, styler.HorizontalAlignment("CENTER").CheckBoxCell(invoice.Paid, &cellBorders))

		row := sheets.RowData{
			Values: cells,
		}
		newRows = append(newRows, &row)
	}

	// !!! THIS IS NOT ZERO INDEXED !!!
	// E.X: If the last column you are updating is column A, the endColumnIndex is 1.
	endColumnIndex := int64(6)

	// !!! THESE ARE ZERO INDEXED !!!
	startColumnIndex := int64(1) // Starting column = B
	startRowIndex := int64(1)    // Starting row = 2

	if err := rwsheets.UpdateSheetData(ssid, endColumnIndex, gid, startColumnIndex, startRowIndex, newRows, srv); err != nil {
		log.Fatalf("failed to update sheet data - %s\n", err.Error())
		return
	}

	log.Println("Successfully updated sheet data!")
}

type SampleData struct {
	Headers  []string   `json:"headers"`
	Invoices []*Invoice `json:"invoices"`
}

type Invoice struct {
	Customer string  `json:"customer"`
	Invoice  string  `json:"invoice"`
	Amount   float64 `json:"amount"`
	Date     string  `json:"date"`
	Paid     bool    `json:"paid"`
}

func getSampleData(file string) (*SampleData, error) {
	// Open the sample data file.
	f, err := os.Open(file)
	if err != nil {
		// Couldn't open the file.
		return nil, err
	}
	defer f.Close()

	// Create a new SampleData.
	data := &SampleData{}

	// Decode the contents of the file into the data object.
	err = json.NewDecoder(f).Decode(data)
	return data, err
}
