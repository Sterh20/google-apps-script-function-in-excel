# Google Apps Script Function in Excel

[![GitHub stars](https://img.shields.io/github/stars/Sterh20/google-apps-script-function-in-excel.svg?style=social&label=Stars)](https://github.com/Sterh20/google-apps-script-function-in-excel/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/Sterh20/google-apps-script-function-in-excel.svg?style=social&label=Forks)](https://github.com/Sterh20/google-apps-script-function-in-excel/network/members)
[![GitHub watchers](https://img.shields.io/github/watchers/Sterh20/google-apps-script-function-in-excel.svg?style=social&label=Watchers)](https://github.com/Sterh20/google-apps-script-function-in-excel/watchers)
[![GitHub followers](https://img.shields.io/github/followers/Sterh20.svg?style=social&label=Followers)](https://github.com/Sterh20/?tab=followers)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

![Microsoft Excel Badge](https://img.shields.io/badge/Microsoft%20Excel-217346?logo=microsoftexcel&logoColor=fff&style=flat)
![Google Apps Script Badge](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=googleappsscript&logoColor=fff&style=flat)
![JavaScript Badge](https://img.shields.io/badge/JavaScript-F7DF1E?logo=javascript&logoColor=000&style=flat)
![Node.js Badge](https://img.shields.io/badge/Node.js-393?logo=nodedotjs&logoColor=fff&style=flat)
![ESLint Badge](https://img.shields.io/badge/ESLint-4B32C3?logo=eslint&logoColor=fff&style=flat)
![Prettier Badge](https://img.shields.io/badge/Prettier-F7B93E?logo=prettier&logoColor=fff&style=flat)

This repository contains a proof of concept of using **Google Apps Script app as a cloud user defined function (UDF) for Excel**. The code demonstrates how to handle incoming requests from Excel, process the data, and generate a response in the form of an XML document readable by Excel.

## Problem statement

Have you ever been in a situation where you couldn't use VBA's UDF function? Whether it's due to your corporate overlord's internal security policy or your clients' demands, it can be frustrating to create multiple rows, sheets, or even scrap everything just to implement a simple VBA function in a cumbersome and inefficient way. Wouldn't it be cool if you could offload all VBA function's calculations to another place and simply get the function's result in a cell?

Look no further. This project implements the basic functionality required to kickstart the replacement of a VBA UDF function in Excel. If your formula doesn't use multidimensional arrays as input, all you have to do is implement your VBA function's logic in JavaScript. As a result, instead of having a VBA function that executes on your hardware, you will have a JavaScript function that is executed on Google's hardware and returns only the results.

## How it works

The easiest way to understand how this solution works is to examine the [**demo**](demo/demo.xlsx) Excel file. Under each cloud formula's row, there is a hidden formula breakdown in an Excel's group.

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

* `main.gs`: Entry point file that contains the `doGet` function, serving as the Cloud Function endpoint. `responseDataWrapper` function contains the main logic for processing requests, as well as helper functions' calls.
* `helpers.gs`: Includes helper functions used in `main.gs`.
* `NumberParser.gs`: Defines the `NumberParser` class for parsing strings and converting them to numbers taking into account the user's locale settings.
* `tests.gs`: Includes functions to test/try `NumberParser` class and `createXMLDocument` function.

## Usage

To use this code as a Cloud Function in Google Apps Script:

1. Create a new Google Apps Script project.
2. Copy the code from the respective files into the corresponding files in your project or clone this repo and follow the instructions in [setup_procedure.md](.setup/setup_procedure.md):

   2.1 From **"Project Settings"** of your Apps Script project save somewhere **Script ID**

   2.2 In the cloned repo's root directory create `.clasp.json` with the following:

    ```JSON
    {
        "scriptId":"Script-ID-from-Project-Settings",
        "rootDir":"C:\\Local\\path\\to\\repo",
    }
    ```

   2.3 Push repo's files to the Apps Script project by executing (don't forget to activate `.env`):

    ```powershell
    clasp push
    ```

3. Deploy the script as a web app:

   3.1 Push "Deploy" button > "New deployment" in the Apps Script editor's top right side.

   3.2 Select type: `Web app`.

   3.3 Add description. Select "execute as" `Me` and "who has access" `Anyone`

   3.4 Deploy the web app and obtain the URL.
4. Use the deployed app's URL as the endpoint for making HTTP GET requests to the Cloud Function.

## Additional Notes

* The code utilizes the `XmlService` class from the Google Apps Script library to create XML documents.
* The `stripBraces` function removes curly braces from strings if present, and the `stringToArray` function converts strings to arrays based on the locale subtag.
* The `localizedStringToNumber` function converts localized strings or arrays of strings to numbers based on the locale subtag.
* The `formatNumberToLocale` function formats numbers to the specified locale and options.
* The `someFunctionDoableInVBA` function performs an operation on arrays and numbers. This function is used for demonstration purposes only.

## License

Distributed under the MIT license. See [LICENSE](/LICENSE) for more information
