function onOpen() {
}
function insertBrokenLink() {
}
function insertLinkToIssue() {
}
function addIssueHeadings() {
}/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/index.ts");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./node_modules/webpack/buildin/global.js":
/*!***********************************!*\
  !*** (webpack)/buildin/global.js ***!
  \***********************************/
/*! no static exports found */
/***/ (function(module, exports) {

var g;

// This works in non-strict mode
g = (function() {
	return this;
})();

try {
	// This works if eval is allowed (see CSP)
	g = g || new Function("return this")();
} catch (e) {
	// This works if the window reference is available
	if (typeof window === "object") g = window;
}

// g can still be undefined, but nothing to do about it...
// We return undefined, instead of nothing here, so it's
// easier to handle this case. if(!global) { ...}

module.exports = g;


/***/ }),

/***/ "./src/add-issue-headings.ts":
/*!***********************************!*\
  !*** ./src/add-issue-headings.ts ***!
  \***********************************/
/*! exports provided: addIssueHeadings */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "addIssueHeadings", function() { return addIssueHeadings; });
function addIssueHeadings() {
    var doc = DocumentApp.getActiveDocument();
    var element = doc.getCursor().getElement();
    var body = doc.getBody();
    var bodyIndex = body.getChildIndex(element);
    var startAt = parseInt(DocumentApp.getUi()
        .prompt('Enter issue number to start at:')
        .getResponseText());
    var issueTemplate = '#%03d';
    if (!isNaN(startAt)) {
        var lastIssueNumber = startAt + 10;
        for (var issueNumber = startAt; issueNumber < lastIssueNumber; issueNumber++) {
            body
                .insertParagraph(bodyIndex, Utilities.formatString(issueTemplate, issueNumber))
                .setHeading(DocumentApp.ParagraphHeading.HEADING3);
            bodyIndex += 1;
        }
    }
}


/***/ }),

/***/ "./src/index.ts":
/*!**********************!*\
  !*** ./src/index.ts ***!
  \**********************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* WEBPACK VAR INJECTION */(function(global) {/* harmony import */ var _insert_broken_link__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./insert-broken-link */ "./src/insert-broken-link.ts");
/* harmony import */ var _insert_link_to_issue__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./insert-link-to-issue */ "./src/insert-link-to-issue.ts");
/* harmony import */ var _add_issue_headings__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./add-issue-headings */ "./src/add-issue-headings.ts");



global.onOpen = onOpen;
global.insertBrokenLink = _insert_broken_link__WEBPACK_IMPORTED_MODULE_0__["insertBrokenLink"];
global.insertLinkToIssue = _insert_link_to_issue__WEBPACK_IMPORTED_MODULE_1__["insertLinkToIssue"];
global.addIssueHeadings = _add_issue_headings__WEBPACK_IMPORTED_MODULE_2__["addIssueHeadings"];
function onOpen() {
    createUtilitiesMenu();
}
function createUtilitiesMenu() {
    DocumentApp.getUi()
        .createMenu('Utilities')
        .addItem('Insert Broken Link Issue', 'insertBrokenLink')
        .addItem('Insert Link To Issue', 'insertLinkToIssue')
        .addItem('Add issue headings', 'addIssueHeadings')
        // .addItem('Bootstrap update issue', 'insertBootstrapIssue') // TODO: finish this feature
        //    .addItem('Test new form', 'newBrokenLinkDialog')
        .addToUi();
}

/* WEBPACK VAR INJECTION */}.call(this, __webpack_require__(/*! ./../node_modules/webpack/buildin/global.js */ "./node_modules/webpack/buildin/global.js")))

/***/ }),

/***/ "./src/insert-broken-link.ts":
/*!***********************************!*\
  !*** ./src/insert-broken-link.ts ***!
  \***********************************/
/*! exports provided: insertBrokenLink */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "insertBrokenLink", function() { return insertBrokenLink; });
/* harmony import */ var _utils__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./utils */ "./src/utils.ts");

function insertBrokenLink() {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    var element = cursor.getElement();
    var linkText = '';
    var ghLineUrl = '';
    var bulletPoint = '';
    var newParagraphs = [];
    if (element.getType() != DocumentApp.ElementType.PARAGRAPH) {
        return;
    }
    linkText = Object(_utils__WEBPACK_IMPORTED_MODULE_0__["promptUserForLinkText"])();
    if (!linkText) {
        return;
    }
    ghLineUrl = Object(_utils__WEBPACK_IMPORTED_MODULE_0__["promptUserForGithubLineURL"])();
    if (!ghLineUrl) {
        return;
    }
    // TODO: prompt user for bullet point annotation
    //  bulletPoint = promptUserForBulletPoint()
    newParagraphs = appendBrokenLinkToParagraph(element, linkText, ghLineUrl);
    Object(_utils__WEBPACK_IMPORTED_MODULE_0__["moveCursorToEndOfElement"])(newParagraphs[newParagraphs.length - 1]);
}
function appendBrokenLinkToParagraph(paragraph, linkText, ghLineUrl) {
    var parent = paragraph.getParent();
    var i = parent.getChildIndex(paragraph);
    return [
        paragraph.appendText(Utilities.formatString(' - Broken link: %s', linkText)),
        parent.insertParagraph(i + 1, ghLineUrl).setLinkUrl(ghLineUrl),
        parent.insertParagraph(i + 2, ''),
        parent.insertParagraph(i + 3, Utilities.formatString('"%s" link redirects to 404 page on GitHub. (GitHub paths are case sensitive.)', linkText)),
        parent.insertParagraph(i + 4, '')
    ];
}


/***/ }),

/***/ "./src/insert-link-to-issue.ts":
/*!*************************************!*\
  !*** ./src/insert-link-to-issue.ts ***!
  \*************************************/
/*! exports provided: insertLinkToIssue */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "insertLinkToIssue", function() { return insertLinkToIssue; });
/* harmony import */ var _utils__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./utils */ "./src/utils.ts");

function insertLinkToIssue() {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    var element = cursor.getElement();
    var linkText = '';
    var ghLineUrl = '';
    var targetPath = '';
    var newParagraphs = [];
    if (element.getType() != DocumentApp.ElementType.PARAGRAPH) {
        return;
    }
    linkText = Object(_utils__WEBPACK_IMPORTED_MODULE_0__["promptUserForLinkText"])();
    if (!linkText) {
        return;
    }
    ghLineUrl = Object(_utils__WEBPACK_IMPORTED_MODULE_0__["promptUserForGithubLineURL"])();
    if (!ghLineUrl) {
        return;
    }
    targetPath = Object(_utils__WEBPACK_IMPORTED_MODULE_0__["promptUserForTargetPath"])();
    if (!targetPath) {
        return;
    }
    newParagraphs = appendLinkToIssueToParagraph(element, linkText, ghLineUrl, targetPath);
    Object(_utils__WEBPACK_IMPORTED_MODULE_0__["moveCursorToEndOfElement"])(newParagraphs[newParagraphs.length - 1]);
}
function appendLinkToIssueToParagraph(paragraph, linkText, ghLineUrl, targetPath) {
    var parent = paragraph.getParent();
    var i = parent.getChildIndex(paragraph);
    return [
        paragraph.appendText(Utilities.formatString(' - Link to %s', linkText)),
        parent.insertParagraph(i + 1, ghLineUrl).setLinkUrl(ghLineUrl),
        parent.insertParagraph(i + 2, ''),
        parent.insertParagraph(i + 3, Utilities.formatString('"%s" should link to %s.', linkText, targetPath.replace(/\\/gi, '/'))),
        parent.insertParagraph(i + 4, '')
    ];
}


/***/ }),

/***/ "./src/utils.ts":
/*!**********************!*\
  !*** ./src/utils.ts ***!
  \**********************/
/*! exports provided: promptUserForGithubLineURL, promptUserForLinkText, promptUserForBulletPoint, moveCursorToEndOfElement, promptUserForTargetPath */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "promptUserForGithubLineURL", function() { return promptUserForGithubLineURL; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "promptUserForLinkText", function() { return promptUserForLinkText; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "promptUserForBulletPoint", function() { return promptUserForBulletPoint; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "moveCursorToEndOfElement", function() { return moveCursorToEndOfElement; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "promptUserForTargetPath", function() { return promptUserForTargetPath; });
function promptUserForGithubLineURL() {
    var ui = DocumentApp.getUi();
    var response = ui.prompt('Enter the url for the GitHub source code line:');
    var result = '';
    if (response.getSelectedButton() == ui.Button.OK) {
        result = response.getResponseText().trim();
    }
    return result;
}
function promptUserForLinkText() {
    var ui = DocumentApp.getUi();
    var response = ui.prompt('Enter the text for the broken link:');
    var result = '';
    if (response.getSelectedButton() == ui.Button.OK) {
        result = response.getResponseText().trim();
    }
    return result;
}
function promptUserForBulletPoint() {
    var ui = DocumentApp.getUi();
    var button = ui.alert('Add a bullet point annotation to link?', ui.ButtonSet.YES_NO);
    return button == ui.Button.YES;
}
function moveCursorToEndOfElement(element) {
    var position = DocumentApp.getActiveDocument().newPosition(element, element.getText().length);
    DocumentApp.getActiveDocument().setCursor(position);
}
function promptUserForTargetPath() {
    var ui = DocumentApp.getUi();
    var response = ui.prompt('Enter the path to the link target:');
    var result = '';
    if (response.getSelectedButton() == ui.Button.OK) {
        result = response.getResponseText().trim();
    }
    return result;
}


/***/ })

/******/ });