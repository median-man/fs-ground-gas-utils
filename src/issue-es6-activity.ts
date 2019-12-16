/* 
  Use this module for creating an issue to refactor an
  activity to use ES6 syntax.
*/
import { insertActivityLink } from './insert-activity-link'

class IssueES6Activity {
  private DocumentApp: GoogleAppsScript.Document.DocumentApp
  private ui: GoogleAppsScript.Base.Ui

  constructor(options: { DocumentApp: GoogleAppsScript.Document.DocumentApp }) {
    this.DocumentApp = options.DocumentApp
    this.ui = options.DocumentApp.getUi()
  }

  execute() {
    const cursor = this.DocumentApp.getActiveDocument().getCursor()
    if (!cursor) {
      return
    }
    const textElement = cursor.insertText(' - : Refactor with ES6 syntax')
    if (!textElement) {
      // can't insert text at cursor
      return
    }

    // move cursor to after the colon
    this.moveCursor(
      textElement,
      textElement.findText(':').getEndOffsetInclusive()
    )

    insertActivityLink()
    const paragraph = this.insertIssueBodyAtPosition(cursor)
    this.moveCursor(paragraph, 1)
  }

  moveCursor(element, offset) {
    const doc = this.DocumentApp.getActiveDocument()
    const newPosition = doc.newPosition(element, offset)
    doc.setCursor(newPosition)
  }

  insertIssueBodyAtPosition(position: GoogleAppsScript.Document.Position) {
    const body = DocumentApp.getActiveDocument().getBody()
    const index = body.getChildIndex(position.getElement()) + 1
    const issueText =
      'Refactor the code in this activity with the ES6 syntax that was covered in previous units. (e.g. let/const, arrow functions, promises, etc…)'
    return body.insertParagraph(index, issueText)
  }
}

export function createES6ActivityIssue() {
  new IssueES6Activity({ DocumentApp }).execute()
}
