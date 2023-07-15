# Google Apps Script Cloud Function

This repository contains a proof of concept of using Google Apps Script as a Cloud Function. The code demonstrates how to handle incoming requests, process the data, and generate a response in the form of an XML document.

## Functionality

The main code file, `doGet.gs`, includes the following key features:

- Parsing and processing the data received in the request
- Formatting numbers based on locale subtags
- Converting localized strings to numbers
- Creating an XML document from the processed data

The code is organized into separate files to improve modularity and maintainability:

- `main.gs`: Entry point file that contains the `doGet` function, serving as the Cloud Function endpoint.
- `doGet.gs`: Contains the main logic for processing requests, as well as helper functions.
- `helpers.gs`: Includes additional helper functions used in `doGet.gs`.
- `NumberParser.gs`: Defines the `NumberParser` class for parsing strings and converting them to numbers.

## Usage

To use this code as a Cloud Function in Google Apps Script:

1. Create a new Google Apps Script project in the Apps Script editor.
2. Copy the code from the respective files into the corresponding files in your project.
3. Deploy the script as a web app:
   - Go to "Publish" > "Deploy as web app" in the Apps Script editor.
   - Set the project version and configure the web app settings.
   - Deploy the web app and obtain the URL.
4. Use the deployed URL as the endpoint for making HTTP GET requests to the Cloud Function.

## Additional Notes

- The code includes test functions (`testNumberParserBasicFunctionality` and `testCreateXMLDocument`) that can be used to verify the basic functionality of the NumberParser class and the createXMLDocument function.
- The code utilizes the XmlService class from the Google Apps Script library to create XML documents.
- The `stripBraces` function removes curly braces from strings if present, and the `stringToArray` function converts strings to arrays based on the locale subtag.
- The `localizedStringToNumber` function converts localized strings or arrays of strings to numbers based on the locale subtag.
- The `formatNumberToLocale` function formats numbers to the specified locale and options.
- The `someFunctionDoableInVBA` function performs an operation on arrays and numbers, similar to functionality in VBA.

## License

This project is licensed under the [MIT License](LICENSE).