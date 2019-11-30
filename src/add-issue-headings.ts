export function addIssueHeadings() {
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
