/* 
  Use this module for creating an issue to refactor an
  activity to use ES6 syntax.
*/
import { insertActivityLink } from './insert-activity-link'
import { CreateIssue } from './create-issue'

class IssueES6Activity extends CreateIssue {
  constructor(options: { DocumentApp: GoogleAppsScript.Document.DocumentApp }) {
    super(options)
  }

  execute() {
    const cursor = this.getCursor()
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

  insertIssueBodyAtPosition(position: GoogleAppsScript.Document.Position) {
    const body = this.document.getBody()
    const index = body.getChildIndex(position.getElement()) + 1
    const issueText =
      'Refactor the code in this activity with the ES6 syntax that was covered in previous units. (e.g. let/const, arrow functions, promises, etc…)'
    return body.insertParagraph(index, issueText)
  }
}

export function createES6ActivityIssue() {
  new IssueES6Activity({ DocumentApp }).execute()
}
