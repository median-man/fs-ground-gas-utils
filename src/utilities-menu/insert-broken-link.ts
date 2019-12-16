import {
  promptUserForLinkText,
  promptUserForGithubLineURL,
  moveCursorToEndOfElement
} from '../utils'

export function insertBrokenLink() {
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
