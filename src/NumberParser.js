/**
 * @fileoverview This file contains the NumberParser class, which is used for parsing strings and converting them to numbers.
 * It includes functions for initializing the NumberParser instance and parsing strings based on locale subtags.
 * @author [Sterh20](https://github.com/Sterh20) et al.
 */

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
