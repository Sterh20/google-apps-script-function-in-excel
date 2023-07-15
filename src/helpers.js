/* eslint-disable multiline-ternary */
/**
 * @fileoverview This file contains various helper functions used in the `main.gs` file.
 * These functions handle specific tasks such as stripping braces from strings,
 * converting strings to arrays based on locale subtags, converting localized strings to numbers,
 * formatting numbers to the specified locale, and creating an XML document.
 * @author [Sterh20](https://github.com/Sterh20)
 */

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
 * Limitations: Only for one dimensional array case.
 * @param {string} str - The input string.
 * @param {string} localeSubtag - The locale subtag for formatting.
 * @return {string[]} The array of strings.
 */
function stringToArray (str, localeSubtag = 'en-US') {
  // Only for one dimensional array case. Two dimensional arrays adds additional level of complexity in terms of locale.
  // E.g. in "en-US" locale rows are separated by ";" and columns by ",", but in "fr-FR" or "fr-CA" it is ":" and ";", accordingly.
  // List of separators for different locales on Windows 10 https://www.sigmdel.ca/michel/program/fpl/list_sep_en.html
  // Their is probably a better way of doing this using Intl.js.
  // TODO: Think of a way to mitigate excessive separation for strings in the Array.
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
 * Limitations: Only for one dimensional array case or a number in string form.
 * Works only with arrays containing exclusively numbers.
 * @param {string|string[]} input - The input string or array of strings.
 * @param {string} localeSubtag - The locale subtag for formatting.
 * @return {number|number[]} The converted number or array of numbers.
 */
function localizedStringToNumber (input, localeSubtag = 'en-US') {
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
