/*
 * Third-party code (docx image module for docxtemplater), see ./index.mjs for provenance and license.
 * Derived from docxtemplater-image-module-free 1.1.1 js/docUtils.js
 * Modified: ES module.
 */
/* eslint-disable */

import Docxtemplater from 'docxtemplater'

var DocUtils = Docxtemplater.DocUtils;
DocUtils.convertPixelsToEmus = function (pixel) {
	return Math.round(pixel * 9525);
};
export default DocUtils;
