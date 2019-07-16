package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"github.com/pkg/errors"
	"golang.org/x/net/context"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
	"io/ioutil"
	"log"
	"net/http"
	"os"
)

// Retrieve a token, saves the token, then returns the generated client.
func getClient(config *oauth2.Config) *http.Client {
	// The file token.json stores the user's access and refresh tokens, and is
	// created automatically when the authorization flow completes for the first
	// time.
	tokFile := "token.json"
	tok, err := tokenFromFile(tokFile)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(tokFile, tok)
	}
	return config.Client(context.Background(), tok)
}

// Request a token from the web, then returns the retrieved token.
func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser then type the "+
		"authorization code: \n%v\n", authURL)

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		log.Fatalf("Unable to read authorization code: %v", err)
	}

	tok, err := config.Exchange(context.TODO(), authCode)
	if err != nil {
		log.Fatalf("Unable to retrieve token from web: %v", err)
	}
	return tok
}

// Retrieves a token from a local file.
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

// Saves a token to a file path.
func saveToken(path string, token *oauth2.Token) {
	fmt.Printf("Saving credential file to: %s\n", path)
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		log.Fatalf("Unable to cache oauth token: %v", err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(token)
}

func Setup() *http.Client {
	b, err := ioutil.ReadFile("credentials.json")
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v", err)
	}

	// If modifying these scopes, delete your previously saved token.json.
	config, err := google.ConfigFromJSON(b, "https://www.googleapis.com/auth/spreadsheets")
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}
	client := getClient(config)
	return client
}

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
