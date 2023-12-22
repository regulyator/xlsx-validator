package main

import (
	"encoding/json"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"os"
	"strings"
)

type Validation struct {
	Dictionaries       map[string][]string `json:"dictionaries"`
	Fields             []Field             `json:"fields"`
	KeyField           int                 `json:"keyField"`
	ErrorMessageColumn string              `json:"errorMessageColumn"`
	SkipHeader         bool                `json:"skipHeader"`
}

type Field struct {
	FieldID   int     `json:"fieldID"`
	Type      string  `json:"type"`
	Storage   string  `json:"storage"`
	Separator *string `json:"separator"`
	Rules     []Rule  `json:"rules"`
}

type Rule struct {
	Type         string  `json:"type"`
	Dictionary   *string `json:"dictionary,omitempty"`
	RefField     *int    `json:"refField,omitempty"`
	ErrorMessage string  `json:"errorMessage"`
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Please specify file name")
		return
	}
	fileName := os.Args[1]

	//read validation
	var validation Validation
	_ = readValidation(&validation)

	//work with xlsx
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	var errorCount = 0

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	headersMap := make(map[int]string)
	for idx, v := range rows[0] {
		headersMap[idx] = v
	}

	for iRow, row := range rows {
		if iRow == 0 && validation.SkipHeader {
			continue
		}

		var rowResult strings.Builder
		for _, fieldVal := range validation.Fields {
			if err != nil {
				fmt.Println("Error converting field key to int")
			}

			for _, rule := range fieldVal.Rules {
				switch rule.Type {
				case "NON_NULL":
					checkNonNull(row, &rule, &fieldVal, &rowResult, headersMap)
				case "IN_DICTIONARY":
					checkInDictionary(row, &rule, &fieldVal, &rowResult, headersMap, validation.Dictionaries)
				case "NOT_IN_FIELD":
					checkNotInField(row, &rule, &fieldVal, &rowResult, headersMap)
				default:
					fmt.Println("Unknown rule type")
				}
			}
		}

		if rowResult.Len() > 0 {
			errorCount++
			err := f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", validation.ErrorMessageColumn, iRow+1),
				fmt.Sprintf("%s: %s.%s\n", headersMap[validation.KeyField], row[validation.KeyField], rowResult.String()),
			)
			if err != nil {
				fmt.Println("Error writing error to output file")
			}

		}
	}

	if err := f.SaveAs(fmt.Sprintf("%svalidation_result_%d.xlsx", fileName, errorCount)); err != nil {
		fmt.Println("Error saving result file")
	}

	fmt.Printf("Validation finished, %d errors found\n", errorCount)
}

func checkNotInField(row []string, rule *Rule, fieldVal *Field, rowResult *strings.Builder, headersMap map[int]string) {
	var notInFieldResult strings.Builder
	inRefField := contains(strings.Split(row[*rule.RefField], *fieldVal.Separator))
	fieldValues := strings.Split(row[fieldVal.FieldID], *fieldVal.Separator)
	for _, val := range fieldValues {
		if inRefField(strings.ToUpper(strings.TrimSpace(val))) && len(strings.TrimSpace(val)) > 0 {
			notInFieldResult.WriteString(val)
			notInFieldResult.WriteString(";")
		}
	}
	if notInFieldResult.Len() > 0 {
		rowResult.WriteString(fmt.Sprintf(" Fields %s and %s %s, error values: %s", headersMap[fieldVal.FieldID], headersMap[*rule.RefField], rule.ErrorMessage, notInFieldResult.String()))
	}
}

func checkInDictionary(row []string, rule *Rule, fieldVal *Field, rowResult *strings.Builder, headersMap map[int]string, dictionaries map[string][]string) {
	var fieldDictionaryResult strings.Builder
	inDictionary := contains(dictionaries[*rule.Dictionary])
	fieldValues := strings.Split(row[fieldVal.FieldID], *fieldVal.Separator)
	for _, val := range fieldValues {
		if !inDictionary(strings.ToUpper(strings.TrimSpace(val))) && len(strings.TrimSpace(val)) > 0 {
			fieldDictionaryResult.WriteString(val)
			fieldDictionaryResult.WriteString(";")
		}
	}
	if fieldDictionaryResult.Len() > 0 {
		rowResult.WriteString(fmt.Sprintf(" Field %s %s, values not in dictionary: %s", headersMap[fieldVal.FieldID], rule.ErrorMessage, fieldDictionaryResult.String()))
	}
}

func checkNonNull(row []string, rule *Rule, fieldVal *Field, rowResult *strings.Builder, headersMap map[int]string) {
	if len(strings.TrimSpace(row[fieldVal.FieldID])) == 0 {
		rowResult.WriteString(fmt.Sprintf(" Field %s %s", headersMap[fieldVal.FieldID], rule.ErrorMessage))
	}
}

func contains(list []string) func(string) bool {
	return func(s string) bool {
		for _, v := range list {
			if v == s {
				return true
			}
		}
		return false
	}
}

func readValidation(validation *Validation) error {
	file, err := os.Open(".validate.json")
	if err != nil {
		fmt.Println("Error opening file:", err)
		return err
	}
	defer file.Close()

	// Read the file's contents
	byteValue, err := io.ReadAll(file)
	if err != nil {
		fmt.Println("Error reading file:", err)
		return err
	}

	// Unmarshal the JSON data into the struct
	err = json.Unmarshal(byteValue, validation)
	if err != nil {
		fmt.Println("Error unmarshalling JSON:", err)
		return err
	}
	return nil
}
