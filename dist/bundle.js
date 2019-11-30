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

/***/ "./src/index.ts":
/*!**********************!*\
  !*** ./src/index.ts ***!
  \**********************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

/* WEBPACK VAR INJECTION */(function(global) {global.onOpen = onOpen;
global.insertBrokenLink = insertBrokenLink;
global.insertLinkToIssue = insertLinkToIssue;
global.addIssueHeadings = addIssueHeadings;
function onOpen() {
    // add utilities menu
    DocumentApp.getUi()
        .createMenu('Utilities')
        .addItem('Insert Broken Link Issue', 'insertBrokenLink')
        .addItem('Insert Link To Issue', 'insertLinkToIssue')
        .addItem('Add issue headings', 'addIssueHeadings')
        // .addItem('Bootstrap update issue', 'insertBootstrapIssue') // TODO: finish this feature
        //    .addItem('Test new form', 'newBrokenLinkDialog')
        .addToUi();
}
// TODO: this is a work in progress 😉
function newBrokenLinkDialog() {
    var html = HtmlService.createHtmlOutputFromFile('broken-link');
    //    .setWidth(300)
    //    .setHeight(250)
    DocumentApp.getUi().showModalDialog(html, 'New Broken Link Issue');
}
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
    linkText = promptUserForLinkText();
    if (!linkText) {
        return;
    }
    ghLineUrl = promptUserForGithubLineURL();
    if (!ghLineUrl) {
        return;
    }
    // TODO: prompt user for bullet point annotation
    //  bulletPoint = promptUserForBulletPoint()
    newParagraphs = appendBrokenLinkToParagraph(element, linkText, ghLineUrl);
    moveCursorToEndOfElement(newParagraphs[newParagraphs.length - 1]);
}
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
function moveCursorToEndOfElement(element) {
    var position = DocumentApp.getActiveDocument().newPosition(element, element.getText().length);
    DocumentApp.getActiveDocument().setCursor(position);
}
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
    linkText = promptUserForLinkText();
    if (!linkText) {
        return;
    }
    ghLineUrl = promptUserForGithubLineURL();
    if (!ghLineUrl) {
        return;
    }
    targetPath = promptUserForTargetPath();
    if (!targetPath) {
        return;
    }
    newParagraphs = appendLinkToIssueToParagraph(element, linkText, ghLineUrl, targetPath);
    moveCursorToEndOfElement(newParagraphs[newParagraphs.length - 1]);
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
function isNaN(value) {
    return typeof value === 'number' && value !== value;
}
function insertBootstrapIssue() {
    var response;
    var doc = DocumentApp.getActiveDocument();
    var element = doc.getCursor().getElement();
    var activityTitle = '';
    var lessonPlanLink = '';
    var ghPermalinkUrl = '';
    var ghLineNumber = '';
    var ghLineLabel = '';
    var shouldGetPermalink = true;
    var lineLinks = [];
    var ui = DocumentApp.getUi();
    if (element.getType() != DocumentApp.ElementType.PARAGRAPH) {
        return;
    }
    // prompt user for lesson plan link
    response = ui.prompt('Enter link to lesson plan section:');
    if (response.getSelectedButton() != ui.Button.OK) {
        return;
    }
    lessonPlanLink = response.getResponseText();
    // prompt user for name of activity
    // while shouldGetPermalink
    // get gh line permalink
    // get text for line permalink
    // calculate line number form permalink
    // add to linkLinks
    // ask user if there is another permalink and set shouldGetPermalink
    // end while
    // TODO: finish pseudocode
}

/* WEBPACK VAR INJECTION */}.call(this, __webpack_require__(/*! ./../node_modules/webpack/buildin/global.js */ "./node_modules/webpack/buildin/global.js")))

/***/ })

/******/ });