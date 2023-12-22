### Util for validating xlsx files

#### Usage

**create in root directory file named [.validate.json](.validate.json) with following structure:**

```json
{
  "fields": [
    {
      "fieldID": 1,
      "type": "STRING",
      "storage": "PLAIN",
      "separator": null,
      "rules": [
        {
          "type": "NON_NULL",
          "errorMessage": "must not be null"
        }
      ]
    },
    {
      "fieldID": 2,
      "type": "STRING",
      "storage": "ARRAY",
      "separator": ";",
      "rules": [
        {
          "type": "IN_DICTIONARY",
          "dictionary": "languages",
          "errorMessage": "should be in dictionary"
        }
      ]
    },
    {
      "fieldID": 3,
      "type": "STRING",
      "storage": "ARRAY",
      "separator": ";",
      "rules": [
        {
          "type": "IN_DICTIONARY",
          "dictionary": "countries",
          "errorMessage": "should be in dictionary"
        },
        {
          "type": "NOT_IN_FIELD",
          "refField": 6,
          "errorMessage": "should not be the same"
        }
      ]
    },
    {
      "fieldID": 4,
      "type": "STRING",
      "storage": "ARRAY",
      "separator": ";",
      "rules": [
        {
          "type": "IN_DICTIONARY",
          "dictionary": "countries",
          "errorMessage": "should be in dictionary"
        }
      ]
    }
  ],
  "skipHeader": true,
  "keyField": 0,
  "errorMessageColumn": "Y",
  "dictionaries": {
    "dictionary1": [
      "value1",
      "value2"
    ],
    "dictionary2": [
      "value1",
      "value2"
    ]
  }
}
```

where:

  - `fields` - array of fields to validate
    - `fieldID` - column number in xlsx file
    - `type` - type of field
    - `storage` - storage type of field
    - `separator` - separator for array fields
  - `rules` - array of rules to validate field
    - `type` - type of rule
    - `errorMessage` - error message to show if rule is not satisfied
    - `dictionary` - name of dictionary to use in rule
    - `refField` - column number of field to use in rule
  - `skipHeader` - skip header row
  - `keyField` - column number of key field
  - `errorMessageColumn` - column to show error message
  - `dictionaries` - object with dictionaries map

rules:

  - `NON_NULL` - field must not be null
  - `IN_DICTIONARY` - field must be in dictionary
  - `NOT_IN_FIELD` - field must not be the same as another field

**also you should as arguments pass path to xlsx file**
