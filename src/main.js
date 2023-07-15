/* eslint-disable multiline-ternary */
/**
 * @fileoverview Entry point and request handling for the Cloud Function.
 * Contains the `doGet` function to process requests and return a response in XML format.
 * @author [Sterh20](https://github.com/Sterh20)
 */

/**
 * Cloud Function endpoint that processes requests and returns a response in XML format.
 * @param {Object} e - The request event object.
 * @return {ContentService.TextOutput} The response content in XML format.
 */
function doGet (e) {
  const requestsData = e.parameter;
  const responseData = responseDataWrapper(requestsData);
  // For a single value or an Array, but with additional processing by formulas and all values need to be converted from text to number in Excel
  // return ContentService.createTextOutput(responseData);
  // For a single value or an Array.
  // Do not need to be converted to number or process in case of an Array (acts like a spilled range).
  // Needs "FILTERXML()" formula to work in Excel.
  // Better in terms of extensibility. Error messages can be stored in separate node from data.
  return ContentService.createTextOutput(createXMLDocument(responseData));
}

/**
 * Wraps the data received in the request with the necessary processing and formatting.
 * @param {Object} requestsData - The data received in the request.
 * @return {string} The formatted response data.
 */
function responseDataWrapper (requestsData) {
  const localeFromRequest = 'locale' in requestsData ? requestsData.locale : undefined;
  let stringWithoutBraces = 'data' in requestsData ? stripBraces(requestsData.data) : undefined;
  const arrayOfStrings = stringToArray(stringWithoutBraces, localeFromRequest);
  const arrayOfNumbersFromRequest = localizedStringToNumber(arrayOfStrings, localeFromRequest);
  stringWithoutBraces = 'index' in requestsData ? stripBraces(requestsData.index) : undefined;
  const numberDataFromRequest = localizedStringToNumber(stringWithoutBraces, localeFromRequest);
  const responseData = someFunctionDoableInVBA(arrayOfNumbersFromRequest, numberDataFromRequest);
  const formatterOptions = { maximumFractionDigits: 20, useGrouping: false }; // maximumFractionDigits: 3 by default, max is 20, excel's max is 15
  const formattedResponseData = formatNumberToLocale(responseData, localeFromRequest, formatterOptions);
  return formattedResponseData;
}

/**
 * Performs some operation on an array and a number, mimicking functionality of VBA.
 * @param {number[]} anArray - The input array of numbers.
 * @param {number} aNumber - The input number.
 * @return {number} The result of the operation.
 */
function someFunctionDoableInVBA (anArray, aNumber) {
  const sumOfArrayValues = anArray.reduce((accumulator, value) => accumulator + value, 0);
  const arraySumMultipliedByNumber = sumOfArrayValues * aNumber;
  return arraySumMultipliedByNumber;
}
