# Google Apps Script Function in Excel

This repository contains a proof of concept of using Google Apps Script as a Custom Cloud Function. The code demonstrates how to handle incoming requests, process the data, and generate a response in the form of an XML document.

## Problem statement

Have you ever been in a situation where you couldn't use VBA's UDF function, whether due to your corporate overlord's internal security policy or your clients' demands? Isn't it frustrating to create multiple rows, sheets or even scrap everything to try to implement a simple VBA function in a cumbersome and inefficient way? Wouldn't it be cool if you could offload all VBA function calculations to somewhere and just get the function's result in a cell?

Look no further. This project implements the basic functionality required to kickstart the replacement of a VBA UDF function in Excel. If your formula is not using multidimensional arrays as an input, you would just have to implement your VBA function's logic in JavaScript. As a result, instead of a VBA function that executes on your hardware, you will get a JavaScript function that is executed from Google's hardware.

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