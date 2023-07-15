/* eslint-disable multiline-ternary */
/**
 * @fileoverview Entry point and request handling for the Cloud Function.
 * Contains the `doGet` function to process requests and return a response in XML format.
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
 * Removes the curly braces from a string if present.
 * @param {string} str - The input string.
 * @return {string} The string without braces.
 */
function stripBraces (str) {
  str = str.trim();
  // If the string does not starts and ends with curly braces, return the original string
  if (str.charAt(0) !== '{' || str.charAt(str.length - 1) !== '}') {
    return str;
  } else {
    return str.slice(1, -1).trim();
  }
}

/**
 * Converts a string to an array based on the locale subtag.
 * @param {string} str - The input string.
 * @param {string} localeSubtag - The locale subtag for formatting.
 * @return {string[]} The array of strings.
 */
function stringToArray (str, localeSubtag = 'en-US') {
  // Only for one dimensional array case. Two dimensional arrays adds additional level of complexity in terms of locale.
  // E.g. in "en-US" locale rows are separated by ";" and columns by ",", but in "fr-FR" or "fr-CA" it is ":" and ";", accordingly.
  // Think of a way to mitigate excessive separation for strings in the Array.
  // List of separators for different locales on Windows 10 https://www.sigmdel.ca/michel/program/fpl/list_sep_en.html
  // Their is probably a better way of doing this using Intl.js.
  if (['en-US', 'es-US', 'en-GB', 'en-AU', 'en-CA',
    'en-IE', 'en-IN', 'en-MY', 'en-SG', 'as-IN',
    'en-ZW', 'es-MX', 'vi-VN', 'he-IL', 'hi-IN',
    'ja-JP', 'zh-TW', 'bn-IN'
  ].includes(localeSubtag)) {
    return str.split(',');
  } else if (['ru-RU', 'uk-UA', 'be-BY', 'ru-MD', 'ro-MD'].includes(localeSubtag)) {
    return str.split(';');
  } else if (['dv-MV'].includes(localeSubtag)) {
    return str.split('?');
  } else if (['ku-Arab-IQ', 'fa-IR'].includes(localeSubtag)) {
    return str.split('º');
  } else if (str.split(';')[0] === str) {
    return str.split(',');
  } else {
    return str.split(';');
  }
}

/**
 * Converts a localized string or an array of strings to numbers based on the locale subtag.
 * @param {string|string[]} input - The input string or array of strings.
 * @param {string} localeSubtag - The locale subtag for formatting.
 * @return {number|number[]} The converted number or array of numbers.
 */
function localizedStringToNumber (input, localeSubtag = 'en-US') {
  // Only for one dimensional array case or a number in string form.
  // Works only with arrays containing exclusively numbers.
  const strToNumberParser = new NumberParser(localeSubtag);
  if (Array.isArray(input)) {
    return input.map(stringNumber => strToNumberParser.parse(stringNumber));
  } else {
    return strToNumberParser.parse(input);
  }
}

/**
 * Formats a number to the specified locale and options.
 * @param {number} num - The input number.
 * @param {string} localeSubtag - The locale subtag for formatting.
 * @param {Object} options - The formatting options.
 * @return {string} The formatted number.
 */
function formatNumberToLocale (num, localeSubtag = 'en-US', options = {}) {
  return new Intl.NumberFormat(localeSubtag, options).format(num);
}

/**
 * Performs some operation on an array and a number, mimicking functionality in VBA.
 * @param {number[]} anArray - The input array of numbers.
 * @param {number} aNumber - The input number.
 * @return {number} The result of the operation.
 */
function someFunctionDoableInVBA (anArray, aNumber) {
  const sumOfArrayValues = anArray.reduce((accumulator, value) => accumulator + value, 0);
  const arraySumMultipliedByNumber = sumOfArrayValues * aNumber;
  return arraySumMultipliedByNumber;
}

/**
 * A utility class for parsing localized number strings and converting them to numbers.
 * Supports handling various number formats based on the provided locale subtag.
 *
 * @class
 * @param {string} localeSubtag - The locale subtag for formatting.
 * @param {Object} options - The formatting options.
 */
class NumberParser {
  // Compilation of code from different resources.
  // Heavily based on Mike Bostock's code: https://observablehq.com/@mbostock/localized-number-parsing#NumberParser
  // Some structural decisions from p.s.w.g's code: https://stackoverflow.com/a/55366435
  // Negative numbers handling from Evgeni Dikerman's code: https://stackoverflow.com/a/74352853
  // Fix for whitespace in groupings by Mike Reinstein: https://observablehq.com/@mbostock/localized-number-parsing#comment-d5e1fbdc82d07f25
  /**
   * Constructs a NumberParser instance.
   * @param {string} localeSubtag - The locale subtag for formatting.
   * @param {Object} options - The formatting options.
   */
  constructor (localeSubtag, options = {}) {
    const numberFormatter = new Intl.NumberFormat(localeSubtag, options); // Options docs: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Intl/NumberFormat/NumberFormat#options
    // Creates standard for the locale group, decimal, integer and fraction from a number.
    const numberParts = numberFormatter.formatToParts(-12345.6);
    // Creates an array of digits that used in this locale.
    const localesNumerals = Array.from({ length: 10 }).map((_, i) => numberFormatter.format(i));
    const numeralsIndexMapper = new Map(localesNumerals.map((d, i) => [d, i]));
    // Fix for whitespace by Mike Reinstein.
    const localesGroupSeparator = numberParts.find(d => d.type === 'group').value;
    const isWhiteSpace = /\s/.test(
      typeof localesGroupSeparator === 'number' ? String.fromCodePoint(localesGroupSeparator) : localesGroupSeparator.charAt(0)
    );
    this._localesMinusSign = numberParts[0].value;
    // this._groupSeparatorRegExp = new RegExp(`[${numberParts.find(d => d.type === "group").value}]`, "g"); // Old solution in case of bugs.
    this._groupSeparatorRegExp = new RegExp(`[${isWhiteSpace ? '\\s' : localesGroupSeparator}]`, 'g');
    this._decimalSeparatorRegExp = new RegExp(`[${numberParts.find(d => d.type === 'decimal').value}]`);
    this._numeralsRegExp = new RegExp(`[${localesNumerals.join('')}]`, 'g');
    this._numeralsIndexMapper = d => numeralsIndexMapper.get(d);
  }

  /**
   * Parses a string and converts it to a number.
   * @param {string} str - The input string.
   * @return {number} The parsed number.
   */
  parse (str) {
    // eslint-disable-next-line no-cond-assign
    return (str = str.trim()
      .replace(this._localesMinusSign, '-')
      .replace(this._groupSeparatorRegExp, '')
      .replace(this._decimalSeparatorRegExp, '.')
      .replace(this._numeralsRegExp, this._numeralsIndexMapper)
    ) ? +str : NaN;
  }
}

/**
 * Tests the basic functionality of the NumberParser class.
 * @return {boolean} Indicates if the test passed.
 */
function testNumberParserBasicFunctionality () {
  const yourLocale = 'en-US'; // Easy example
  // const yourLocale = "ar-EG"; // Base example
  // const yourLocale = "fr-CA"; // Mike Reinstein example
  const formatter = new Intl.NumberFormat(yourLocale);
  const testNumber = -1234.5; // For "en-US" and "ar-EG"
  const formattedToLocaleNumber = formatter.format(testNumber); // For "en-US" and "ar-EG"
  // const formattedToLocaleNumber = "1 999,99"; // Mike Reinstein example
  const parser = new NumberParser(yourLocale);
  console.log(formattedToLocaleNumber); // ١٬٢٣٤٫٥ in "ar-EG"
  console.log(parser.parse(formattedToLocaleNumber)); // 1234.5
  return true;
}

/**
 * Creates an XML document from the input data.
 * @param {string|string[]} input - The input data.
 * @return {string} The XML document as a string.
 */
function createXMLDocument (input) {
  // Create a new XML document
  const xmlDocument = XmlService.createDocument();

  // Create the root element
  const rootElement = XmlService.createElement('root');
  xmlDocument.setRootElement(rootElement);

  // Function to create XML elements recursively
  function createXMLElement (parentElement, value) {
    if (Array.isArray(value)) {
      // If the value is an array, create child elements for each item
      for (let i = 0; i < value.length; i++) {
        createXMLElement(parentElement, value[i]);
      }
    } else {
      // If the value is a string or number, create a child element with the value
      const childElement = XmlService.createElement('item');
      childElement.setText(value);
      parentElement.addContent(childElement);
    }
  }

  // Call the function to create XML elements based on the input
  createXMLElement(rootElement, input);

  // Convert the XML document to a string
  const xmlString = XmlService.getRawFormat().format(xmlDocument);

  // Return the XML string
  return xmlString;
}

/**
 * Tests the createXMLDocument function.
 */
function testCreateXMLDocument () {
  const input1 = 'Hello';
  const xmlString1 = createXMLDocument(input1);
  console.log(xmlString1);

  const input2 = [1, 2, 3, 4, 5];
  const xmlString2 = createXMLDocument(input2);
  console.log(xmlString2);
}
