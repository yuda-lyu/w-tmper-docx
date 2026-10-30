/*
 * Image module for docxtemplater (docx only), vendored into w-tmper-docx.
 *
 * Derived from docxtemplater-image-module-free 1.1.1 (js/index.js, babel compiled output)
 *   author: Dk Saha, license: MIT, repository: https://github.com/evilc0des/docxtemplater-image-module-free
 * which is a free edition of docxtemplater-image-module
 *   Copyright (c) 2013 Edgar HIPP, dual licensed under the MIT license or the GPLv3 (used here under the MIT license)
 *
 * The MIT License
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
 * documentation files (the "Software"), to deal in the Software without restriction, including without limitation
 * the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
 * and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions
 * of the Software. THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND. The full license text is in
 * the LICENSE file at the root of the w-tmper-docx package.
 *
 * Modifications for w-tmper-docx:
 *   - CommonJS converted to ES modules (index.mjs, imgManager.mjs, templates.mjs, docUtils.mjs).
 *   - Removed the pptx path (slide parsing, pptx picture rendering and template, ppt prefix of imgManager) and with it
 *     the xmldom dependency, which was only used to parse pptx slides. Only docx is supported.
 *   - optionsTransformer throws when the template is not docx (instead of silently using the docx path).
 */
/* eslint-disable */

import Docxtemplater from 'docxtemplater'
import templates from './templates.mjs'
import ImgManager from './imgManager.mjs'

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var DocUtils = Docxtemplater.DocUtils;

function isNaN(number) {
	return !(number === number);
}

var moduleName = "open-xml-templating/docxtemplater-image-module";

function getInnerDocx(_ref) {
	var part = _ref.part;

	return part;
}

var ImageModule = function () {
	function ImageModule(options) {
		_classCallCheck(this, ImageModule);

		this.name = "ImageModule";
		this.options = options || {};
		this.imgManagers = {};
		if (this.options.centered == null) {
			this.options.centered = false;
		}
		if (this.options.getImage == null) {
			throw new Error("You should pass getImage");
		}
		if (this.options.getSize == null) {
			throw new Error("You should pass getSize");
		}
		this.imageNumber = 1;
	}

	_createClass(ImageModule, [{
		key: "optionsTransformer",
		value: function optionsTransformer(options, docxtemplater) {
			var relsFiles = docxtemplater.zip.file(/\.xml\.rels/).concat(docxtemplater.zip.file(/\[Content_Types\].xml/)).map(function (file) {
				return file.name;
			});
			this.fileTypeConfig = docxtemplater.fileTypeConfig;
			this.fileType = docxtemplater.fileType;
			if (this.fileType !== "docx") {
				throw new Error("ImageModule only supports docx templates, got fileType: " + this.fileType);
			}
			this.zip = docxtemplater.zip;
			options.xmlFileNames = options.xmlFileNames.concat(relsFiles);
			return options;
		}
	}, {
		key: "set",
		value: function set(options) {
			if (options.zip) {
				this.zip = options.zip;
			}
			if (options.xmlDocuments) {
				this.xmlDocuments = options.xmlDocuments;
			}
		}
	}, {
		key: "parse",
		value: function parse(placeHolderContent) {
			var module = moduleName;
			var type = "placeholder";
			if (this.options.setParser) {
				return this.options.setParser(placeHolderContent);
			}
			if (placeHolderContent.substring(0, 2) === "%%") {
				return { type: type, value: placeHolderContent.substr(2), module: module, centered: true };
			}
			if (placeHolderContent.substring(0, 1) === "%") {
				return { type: type, value: placeHolderContent.substr(1), module: module, centered: false };
			}
			return null;
		}
	}, {
		key: "postparse",
		value: function postparse(parsed) {
			var expandTo = this.options.centered ? "w:p" : "w:t";
			var getInner = getInnerDocx;
			return DocUtils.traits.expandToOne(parsed, { moduleName: moduleName, getInner: getInner, expandTo: expandTo });
		}
	}, {
		key: "render",
		value: function render(part, options) {
			if (!part.type === "placeholder" || part.module !== moduleName) {
				return null;
			}
			var tagValue = options.scopeManager.getValue(part.value, {
				part: part
			});
			if (!tagValue) {
				return { value: this.fileTypeConfig.tagTextXml };
			} else if ((typeof tagValue === "undefined" ? "undefined" : _typeof(tagValue)) === "object") {
				return this.getRenderedPart(part, tagValue.rId, tagValue.sizePixel);
			}

			var imgManager = new ImgManager(this.zip, options.filePath, this.xmlDocuments);
			var imgBuffer = this.options.getImage(tagValue, part.value);
			var rId = imgManager.addImageRels(this.getNextImageName(), imgBuffer);
			var sizePixel = this.options.getSize(imgBuffer, tagValue, part.value);
			return this.getRenderedPart(part, rId, sizePixel);
		}
	}, {
		key: "resolve",
		value: function resolve(part, options) {
			var _this = this;

			var imgManager = new ImgManager(this.zip, options.filePath, this.xmlDocuments);
			if (!part.type === "placeholder" || part.module !== moduleName) {
				return null;
			}
			var value = options.scopeManager.getValue(part.value, {
				part: part
			});
			if (!value) {
				return { value: this.fileTypeConfig.tagTextXml };
			}
			return new Promise(function (resolve) {
				var imgBuffer = _this.options.getImage(value, part.value);
				resolve(imgBuffer);
			}).then(function (imgBuffer) {
				var rId = imgManager.addImageRels(_this.getNextImageName(), imgBuffer);
				return new Promise(function (resolve) {
					var sizePixel = _this.options.getSize(imgBuffer, value, part.value);
					resolve(sizePixel);
				}).then(function (sizePixel) {
					return {
						rId: rId,
						sizePixel: sizePixel
					};
				});
			});
		}
	}, {
		key: "getRenderedPart",
		value: function getRenderedPart(part, rId, sizePixel) {
			if (isNaN(rId)) {
				throw new Error("rId is NaN, aborting");
			}
			var size = [DocUtils.convertPixelsToEmus(sizePixel[0]), DocUtils.convertPixelsToEmus(sizePixel[1])];
			var centered = this.options.centered || part.centered;
			var newText = this.getRenderedPartDocx(rId, size, centered);
			return { value: newText };
		}
	}, {
		key: "getRenderedPartDocx",
		value: function getRenderedPartDocx(rId, size, centered) {
			return centered ? templates.getImageXmlCentered(rId, size) : templates.getImageXml(rId, size);
		}
	}, {
		key: "getNextImageName",
		value: function getNextImageName() {
			var name = "image_generated_" + this.imageNumber + ".png";
			this.imageNumber++;
			return name;
		}
	}]);

	return ImageModule;
}();

export default ImageModule;
