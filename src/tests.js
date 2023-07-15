/**
 * @fileoverview Test suite for the utility helpers functions.
 * @author [Sterh20](https://github.com/Sterh20)
 */

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
