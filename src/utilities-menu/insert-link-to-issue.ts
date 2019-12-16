import {
  moveCursorToEndOfElement,
  promptUserForLinkText,
  promptUserForGithubLineURL,
  promptUserForTargetPath
} from '../utils'

export function insertLinkToIssue() {
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
