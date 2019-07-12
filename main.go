package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"github.com/pkg/errors"
	"golang.org/x/net/context"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
	"log"
)

func readFromSource(srv *sheets.Service, spreadsheetId, readRange string) (map[string][]string, error) {
	data := make(map[string][]string)
	resp, err := srv.Spreadsheets.Values.Get(spreadsheetId, readRange).Do()
	if err != nil {
		return nil, errors.Wrap(err, "Unable to retrieve data from sheet:")
	}

	if len(resp.Values) == 0 {
		return nil, errors.Wrap(err, "No data found.")
	}

	j, err := json.MarshalIndent(resp.Values, "", "\t")
	if err != nil {
		return nil, err
	}

	temp := make([][]string, 0)
	err = json.Unmarshal(j, &temp)
	if err != nil {
		return nil, err
	}

	for _, t := range temp {
		data[t[0]] = t[1:]
	}

	return data, nil
}

func readList(srv *sheets.Service, spreadsheetId, readRange string) ([]string, error) {
	data := make([]string, 0)
	resp, err := srv.Spreadsheets.Values.Get(spreadsheetId, readRange).Do()
	if err != nil {
		return nil, errors.Wrap(err, "Unable to retrieve data from sheet:")
	}

	if len(resp.Values) == 0 {
		return nil, errors.Wrap(err, "No data found.")
	}

	j, err := json.MarshalIndent(resp.Values, "", "\t")
	if err != nil {
		return nil, err
	}

	temp := make([][]string, 0)
	err = json.Unmarshal(j, &temp)
	if err != nil {
		return nil, err
	}

	for _, t := range temp {
		data = append(data, t[0])
	}

	return data, nil
}

func fillData(list []string, data map[string][]string) []*sheets.RowData {

	rowData := make([]*sheets.RowData, 0)

	for _, student := range list {
		cellData := make([]*sheets.CellData, 0)

		for _, studentData := range data[student] {
			cell := &sheets.CellData{
				UserEnteredValue: &sheets.ExtendedValue{
					StringValue: studentData,
				},
			}
			cellData = append(cellData, cell)
		}

		row := &sheets.RowData{
			Values: cellData,
		}
		rowData = append(rowData, row)
	}

	return rowData
}

func main() {
	client := Setup()
	ctx := context.Background()

	srv, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	sourceSpreadsheetId, sourceReadRange, listSpreadsheetId, listSpreadsheetRange := "", "", "", ""

	flag.StringVar(&sourceSpreadsheetId, "sourceSheet", "", "Spreadsheet id of database.")
	flag.StringVar(&listSpreadsheetId, "listSheet", "", "Spreadsheet id of list.")
	flag.StringVar(&sourceReadRange, "sourceRange", "", "Range of database.")
	flag.StringVar(&listSpreadsheetRange, "listRange", "", "Range of list.")
	flag.Parse()

	// Fetches Data from source:
	data, err := readFromSource(srv, sourceSpreadsheetId, sourceReadRange)
	if err != nil {
		fmt.Println(err)
	}

	// Fetches Enrollment Number List:
	list, err := readList(srv, listSpreadsheetId, listSpreadsheetRange)
	if err != nil {
		fmt.Println(err)
	}

	// Adds Data.
	rowData := fillData(list, data)
	_, err = srv.Spreadsheets.BatchUpdate(listSpreadsheetId, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			{
				AppendDimension: &sheets.AppendDimensionRequest{
					Dimension: "COLUMNS",
					Length: 92,
				},
			},
			{
				UpdateCells: &sheets.UpdateCellsRequest{
					Fields: "*",
					Range: &sheets.GridRange{
						StartRowIndex: 0,
						StartColumnIndex: 1,
					},
					Rows: rowData,
				},
			},
		},
	}).Do()
	if err != nil {
		fmt.Println(err)
	}
}
