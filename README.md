# go-rwsheets

## Description

This is a Golang library to retrieve Google Sheet data and update Google sheet data, with additional helper functions for updating the sheet values.

I create a lot of applications using Google Sheets and these are some functions I've found myself commonly using and I thought they might be useful for others.

## Features

- Retrieve a Sheets row data
- Update a Sheet with new row data
- Remove a row from a Sheet
- Create `*sheets.ExtendedValue` variables for various data types
- Convert a string date to a Google Sheets serial date

<details>
    <summary>Click for a fun fact on Serial Dates!</summary>

    Calculating serial dates varies between Google Sheets and Excel!
    Google Sheets uses `12/30/1899` for the start date while Excel uses `1/1/1900`
</details>

## Methods

| Method | Description  |
| :----  | :---------   |
| GetSheetData(ssid, readRange string, srv *sheets.Service) ([]*sheets.RowData, error) | Retrieve the row data from spreadsheet |
| UpdateSheetData(ssid string, endColumnIndex, gid, startColumnIndex, startRowIndex int64, newVals []*sheets.RowData, srv*sheets.Service) error | Update the row data in a spreadsheet |
| RemoveRow(rows []*sheets.RowData, rmvIdx int) []*sheets.RowData | Remove a specific row in a Sheet |
| SerialDate(value, format string) (float64, error) | Convert a string date to a Google Sheet serial date |
| BoolValue(value bool) *sheets.ExtendedValue | For updating UserEnteredValue with a boolean value |
| FormulaValue(value string) *sheets.ExtendedValue | For updating UserEnteredValue with a formula value |
| NumberValue(value float64) *sheets.ExtendedValue | For updating UserEnteredValue with a number value |
| TextValue(value string) *sheets.ExtendedValue | For updating UserEnteredValue with a text value |

### Note

NumberValue should be used in conjuction with SerialDate for updating date values in a Google Sheet.

## Sample Using oAuth2

```go
package main

import (
    "context"
    "encoding/json"
    "fmt"
    "io/ioutil"
    "log"
    "os"

    rwsheets "github.com/cryliss/go-rwsheets"

    "golang.org/x/oauth2"
    "golang.org/x/oauth2/google"
    "google.golang.org/api/option"
    sheets "google.golang.org/api/sheets/v4"
)

func getToken(config *oauth2.Config) *oauth2.Token {
    tokFile := "token.json"
    tok, err := tokenFromFile(tokFile)
    if err != nil {
        tok = getTokenFromWeb(config)
        saveToken(tokFile, tok)
    }
    return tok
}

func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
    authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
    fmt.Printf("Go to the following link in your browser then type the "+
    "authorization code: \n%v\n", authURL)

    var authCode string
    if _, err := fmt.Scan(&authCode); err != nil {
        log.Fatalf("Unable to read authorization code %v", err)
    }

    tok, err := config.Exchange(context.TODO(), authCode)
    if err != nil {
        log.Fatalf("Unable to retrieve token from web %v", err)
    }
    return tok
}

func tokenFromFile(file string) (*oauth2.Token, error) {
    f, err := os.Open(file)
    if err != nil {
        return nil, err
    }
    defer f.Close()
    tok := &oauth2.Token{}
    err = json.NewDecoder(f).Decode(tok)
    return tok, err
}

func saveToken(path string, token *oauth2.Token) {
    fmt.Printf("Saving credential file to: %s\n", path)
    f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
    if err != nil {
        log.Fatalf("Unable to cache oauth token: %v", err)
    }
    defer f.Close()
    json.NewEncoder(f).Encode(token)
}

func getService() *sheets.Service {
    ctx := context.Background()
    b, err := ioutil.ReadFile("credentials.json")
    if err != nil {
        log.Fatalf("Unable to read client secret file: %v", err)
    }

    // If modifying these scopes, delete your previously saved token.json.
    config, err := google.ConfigFromJSON(b, "https://www.googleapis.com/auth/spreadsheets")
    if err != nil {
        log.Fatalf("Unable to parse client secret file to config: %v", err)
    }
    tok := getToken(config)

    srv, err := sheets.NewService(ctx, option.WithTokenSource(config.TokenSource(ctx, tok)))
    if err != nil {
        log.Fatalf("failed to create sheets service")
    }
    return srv
}

// Reads a data from a spreadsheet and updates the second row
// with values of various formats
func main() {
    var err error
    var rows, newRows []*sheets.RowData
    var updatedValues sheets.RowData
    var values []*sheets.CellData

    ssid := "13yQwByikICWeVWP-erVchQJhakiRQSAZNz4dvrDNWRk" // Please set here
    gid := int64(0)                                        // Please set here
    readRange := "Sheet1!A1:D5"                            // Please set here.

    // !!! THIS IS NOT ZERO INDEXED !!!
    // E.X: If the last column you are updating is column A, the endColumnIndex is 1.
    endColumnIndex := int64(4) // Please set here

    // !!! THESE ARE ZERO INDEXED !!!
    startColumnIndex := int64(0) // Please set here
    startRowIndex := int64(1)    // Please set here

    date := "3/2/2023"
    layout := "1/2/2006"
    serialDate, err := rwsheets.SerialDate(date, layout)
    if err != nil {
        fmt.Printf("failed to convert date %s to serial date\n", date)
    }
    formula := `=IF($A2+365<TODAY(), "Viewing 1yr after repo was published", "")`
    textValue := "Hello friend!"
    boolValue := false

    srv := getService()

    // Read the spreadsheet data
    rows, err = rwsheets.GetSheetData(ssid, readRange, srv)
    if err != nil {
        fmt.Printf("failed to get sheet data - %s\n", err.Error())
        return
    }

    fmt.Printf("read %d rows from %s\n", len(rows), readRange)
    fmt.Println(rows)

    // Get the values in the second row
    values = rows[1].Values

    // Update column A with our serial date
    values[0].UserEnteredValue = rwsheets.NumberValue(serialDate)

    // Update column B with our formula
    values[1].UserEnteredValue = rwsheets.FormulaValue(formula)

    // Update column C with our text value
    values[2].UserEnteredValue = rwsheets.TextValue(textValue)

    // Update column D with our bool value
    values[3].UserEnteredValue = rwsheets.BoolValue(boolValue)

    updatedValues.Values = values
    newRows = append(newRows, &updatedValues)

    if err := rwsheets.UpdateSheetData(ssid, endColumnIndex, gid, startColumnIndex, startRowIndex, newRows, srv); err != nil {
        fmt.Printf("failed to update sheet data - %s\n", err.Error())
        return
    }

    fmt.Println("Successfully updated sheet data!")
}
```
