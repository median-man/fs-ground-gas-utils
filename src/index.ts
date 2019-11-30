declare var global: any;
global.onOpen = onOpen
global.insertBrokenLink = insertBrokenLink
global.insertLinkToIssue = insertLinkToIssue
global.addIssueHeadings = addIssueHeadings

function onOpen() {
  // add utilities menu
  DocumentApp.getUi()
    .createMenu('Utilities')
    .addItem('Insert Broken Link Issue', 'insertBrokenLink')
    .addItem('Insert Link To Issue', 'insertLinkToIssue')
    .addItem('Add issue headings', 'addIssueHeadings')
    // .addItem('Bootstrap update issue', 'insertBootstrapIssue') // TODO: finish this feature
    //    .addItem('Test new form', 'newBrokenLinkDialog')
    .addToUi()
}

// TODO: this is a work in progress 😉
function newBrokenLinkDialog() {
  var html = HtmlService.createHtmlOutputFromFile('broken-link')
  //    .setWidth(300)
  //    .setHeight(250)
  DocumentApp.getUi().showModalDialog(html, 'New Broken Link Issue')
}

function insertBrokenLink() {
  var cursor = DocumentApp.getActiveDocument().getCursor()
  var element = cursor.getElement()
  var linkText = ''
  var ghLineUrl = ''
  var bulletPoint = ''
  var newParagraphs = []

  if (element.getType() != DocumentApp.ElementType.PARAGRAPH) {
    return
  }

  linkText = promptUserForLinkText()
  if (!linkText) {
    return
  }
  ghLineUrl = promptUserForGithubLineURL()
  if (!ghLineUrl) {
    return
  }

  // TODO: prompt user for bullet point annotation
  //  bulletPoint = promptUserForBulletPoint()

  newParagraphs = appendBrokenLinkToParagraph(element, linkText, ghLineUrl)
  moveCursorToEndOfElement(newParagraphs[newParagraphs.length - 1])
}

function promptUserForGithubLineURL() {
  var ui = DocumentApp.getUi()
  var response = ui.prompt('Enter the url for the GitHub source code line:')
  var result = ''
  if (response.getSelectedButton() == ui.Button.OK) {
    result = response.getResponseText().trim()
  }
  return result
}

function promptUserForLinkText() {
  var ui = DocumentApp.getUi()
  var response = ui.prompt('Enter the text for the broken link:')
  var result = ''
  if (response.getSelectedButton() == ui.Button.OK) {
    result = response.getResponseText().trim()
  }
  return result
}

function promptUserForBulletPoint() {
  var ui = DocumentApp.getUi()
  var button = ui.alert(
    'Add a bullet point annotation to link?',
    ui.ButtonSet.YES_NO
  )
  return button == ui.Button.YES
}

function appendBrokenLinkToParagraph(paragraph, linkText, ghLineUrl) {
  var parent = paragraph.getParent()
  var i = parent.getChildIndex(paragraph)
  return [
    paragraph.appendText(
      Utilities.formatString(' - Broken link: %s', linkText)
    ),
    parent.insertParagraph(i + 1, ghLineUrl).setLinkUrl(ghLineUrl),
    parent.insertParagraph(i + 2, ''),
    parent.insertParagraph(
      i + 3,
      Utilities.formatString(
        '"%s" link redirects to 404 page on GitHub. (GitHub paths are case sensitive.)',
        linkText
      )
    ),
    parent.insertParagraph(i + 4, '')
  ]
}

function moveCursorToEndOfElement(element) {
  var position = DocumentApp.getActiveDocument().newPosition(
    element,
    element.getText().length
  )
  DocumentApp.getActiveDocument().setCursor(position)
}

function insertLinkToIssue() {
  var cursor = DocumentApp.getActiveDocument().getCursor()
  var element = cursor.getElement()
  var linkText = ''
  var ghLineUrl = ''
  var targetPath = ''
  var newParagraphs = []

  if (element.getType() != DocumentApp.ElementType.PARAGRAPH) {
    return
  }
  linkText = promptUserForLinkText()
  if (!linkText) {
    return
  }
  ghLineUrl = promptUserForGithubLineURL()
  if (!ghLineUrl) {
    return
  }
  targetPath = promptUserForTargetPath()
  if (!targetPath) {
    return
  }
  newParagraphs = appendLinkToIssueToParagraph(
    element,
    linkText,
    ghLineUrl,
    targetPath
  )
  moveCursorToEndOfElement(newParagraphs[newParagraphs.length - 1])
}

function promptUserForTargetPath() {
  var ui = DocumentApp.getUi()
  var response = ui.prompt('Enter the path to the link target:')
  var result = ''
  if (response.getSelectedButton() == ui.Button.OK) {
    result = response.getResponseText().trim()
  }
  return result
}

function appendLinkToIssueToParagraph(
  paragraph,
  linkText,
  ghLineUrl,
  targetPath
) {
  var parent = paragraph.getParent()
  var i = parent.getChildIndex(paragraph)
  return [
    paragraph.appendText(Utilities.formatString(' - Link to %s', linkText)),
    parent.insertParagraph(i + 1, ghLineUrl).setLinkUrl(ghLineUrl),
    parent.insertParagraph(i + 2, ''),
    parent.insertParagraph(
      i + 3,
      Utilities.formatString(
        '"%s" should link to %s.',
        linkText,
        targetPath.replace(/\\/gi, '/')
      )
    ),
    parent.insertParagraph(i + 4, '')
  ]
}

function addIssueHeadings() {
  var doc = DocumentApp.getActiveDocument()
  var element = doc.getCursor().getElement()
  var body = doc.getBody()
  var bodyIndex = body.getChildIndex(element)
  var startAt = parseInt(
    DocumentApp.getUi()
      .prompt('Enter issue number to start at:')
      .getResponseText()
  )
  var issueTemplate = '#%03d'
  if (!isNaN(startAt)) {
    var lastIssueNumber = startAt + 10
    for (
      var issueNumber = startAt;
      issueNumber < lastIssueNumber;
      issueNumber++
    ) {
      body
        .insertParagraph(
          bodyIndex,
          Utilities.formatString(issueTemplate, issueNumber)
        )
        .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      bodyIndex += 1
    }
  }
}

function isNaN(value) {
  return typeof value === 'number' && value !== value
}

function insertBootstrapIssue() {
  var response
  var doc = DocumentApp.getActiveDocument()
  var element = doc.getCursor().getElement()
  var activityTitle = ''
  var lessonPlanLink = ''
  var ghPermalinkUrl = ''
  var ghLineNumber = ''
  var ghLineLabel = ''
  var shouldGetPermalink = true
  var lineLinks = []

  var ui = DocumentApp.getUi()

  if (element.getType() != DocumentApp.ElementType.PARAGRAPH) {
    return
  }

  // prompt user for lesson plan link
  response = ui.prompt('Enter link to lesson plan section:')
  if (response.getSelectedButton() != ui.Button.OK) {
    return
  }
  lessonPlanLink = response.getResponseText()

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
