class InsertActivityLink {
  private DocumentApp: GoogleAppsScript.Document.DocumentApp
  private ui: GoogleAppsScript.Base.Ui

  constructor(options: { DocumentApp: GoogleAppsScript.Document.DocumentApp }) {
    this.DocumentApp = options.DocumentApp
    this.ui = options.DocumentApp.getUi()
  }

  execute() {
    const activityUrl = this.promptUserForActivityUrl()
    if (activityUrl.length) {
      const linkText = this.activityTextFromUrl(activityUrl)
      this.insertLinkAtCursor(activityUrl, linkText)
    }
  }

  promptUserForActivityUrl(): string {
    const promptTitle = 'Insert Activity Link'
    const promptMessage = 'Activity URL (GitHub)'
    const buttons = this.ui.ButtonSet.OK_CANCEL
    const response = this.ui.prompt(promptTitle, promptMessage, buttons)
    let result = ''
    if (response.getSelectedButton() === this.ui.Button.OK) {
      return response.getResponseText().trim()
    }
    return result
  }

  activityTextFromUrl(url: string): string {
    // activity name is the last name in the url path
    return url.match(/\/([^\/]+?)$/)[1] || ''
  }

  insertLinkAtCursor(url: string, text: string) {
    const cursor = this.DocumentApp.getActiveDocument().getCursor()
    if (!cursor) {
      return
    }
    const textElement = cursor.insertText(text)
    if (!textElement) {
      // can't insert text at cursor
      return
    }
    const linkRange = textElement.findText(text)
    textElement.setLinkUrl(
      linkRange.getStartOffset(),
      linkRange.getEndOffsetInclusive(),
      url
    )

    this.moveCursor(textElement, linkRange.getEndOffsetInclusive() + 1)
  }

  moveCursor(element, offset) {
    const doc = this.DocumentApp.getActiveDocument()
    const newPosition = doc.newPosition(element, offset)
    doc.setCursor(newPosition)
  }
}

export function insertActivityLink() {
  new InsertActivityLink({ DocumentApp }).execute()
}
