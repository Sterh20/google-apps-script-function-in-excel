# Google Apps Script Function in Excel

This repository contains a proof of concept of using Google Apps Script as a Custom Cloud Function. The code demonstrates how to handle incoming requests, process the data, and generate a response in the form of an XML document.

## Problem statement

Have you ever been in a situation where you couldn't use VBA's UDF function? Whether it's due to your corporate overlord's internal security policy or your clients' demands, it can be frustrating to create multiple rows, sheets, or even scrap everything just to implement a simple VBA function in a cumbersome and inefficient way. Wouldn't it be cool if you could offload all VBA function's calculations to another place and simply get the function's result in a cell?

Look no further. This project implements the basic functionality required to kickstart the replacement of a VBA UDF function in Excel. If your formula doesn't use multidimensional arrays as input, all you have to do is implement your VBA function's logic in JavaScript. As a result, instead of having a VBA function that executes on your hardware, you will have a JavaScript function that is executed on Google's hardware and returns only the results.

## How it works

The easiest way to understand how this solution works is to examine the [demo](demo/demo.xlsx) Excel file. Under each cloud formula's row, there is a hidden formula breakdown in an Excel's group.

1. The Apps Script app URL and formula's inputs are combined into a single API-like URL.
2. This URL is used as an input for Excel's `WEBSERVICE` function, which has been available in Excel since Office 2010. The function makes a GET request to the Apps Script app.
3. In the Apps Script app, the function's inputs are converted to the appropriate JavaScript data types. The data type transformations take into account the user's locale settings to ensure compatibility with different languages.
4. After the conversion, the data is processed by a function that we are trying to replace in Excel. In the demo, the function is intentionally kept very simple.
5. The function's results are converted back to the user's locale. Then, the results are converted to XML format and sent back to Excel.
6. The XML result from the `WEBSERVICE` Excel function is filtered using the `FILTERXML` Excel function, and the final result is displayed in the cell. XML was chosen because in the case of plain text, the GET response would need to be converted from text to a number format. Additionally, a plain text response limits the payload to only one item without additional processing.

## Functionality

The main code file, `main.gs`, includes the following key features:

* Parsing and processing the data received in the request
* Formatting numbers based on locale subtags
* Converting localized strings to numbers
* Creating an XML document from the processed data

The code is organized into:

* `main.gs`: Entry point file that contains the `doGet` function, serving as the Cloud Function endpoint.
* `doGet.gs`: Contains the main logic for processing requests, as well as helper functions.
* `helpers.gs`: Includes additional helper functions used in `doGet.gs`.
* `NumberParser.gs`: Defines the `NumberParser` class for parsing strings and converting them to numbers.

## Usage

To use this code as a Cloud Function in Google Apps Script:

1. Create a new Google Apps Script project in the Apps Script editor.
2. Copy the code from the respective files into the corresponding files in your project.
3. Deploy the script as a web app:
   * Go to "Publish" > "Deploy as web app" in the Apps Script editor.
   * Set the project version and configure the web app settings.
   * Deploy the web app and obtain the URL.
4. Use the deployed URL as the endpoint for making HTTP GET requests to the Cloud Function.

## Additional Notes

* The code includes test functions (`testNumberParserBasicFunctionality` and `testCreateXMLDocument`) that can be used to verify the basic functionality of the NumberParser class and the createXMLDocument function.
* The code utilizes the XmlService class from the Google Apps Script library to create XML documents.
* The `stripBraces` function removes curly braces from strings if present, and the `stringToArray` function converts strings to arrays based on the locale subtag.
* The `localizedStringToNumber` function converts localized strings or arrays of strings to numbers based on the locale subtag.
* The `formatNumberToLocale` function formats numbers to the specified locale and options.
* The `someFunctionDoableInVBA` function performs an operation on arrays and numbers, similar to functionality in VBA.

## License

This project is licensed under the [MIT License](LICENSE).
