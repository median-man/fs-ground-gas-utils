export abstract class CreateIssue {
  protected DocumentApp: GoogleAppsScript.Document.DocumentApp
  protected ui: GoogleAppsScript.Base.Ui
  protected document: GoogleAppsScript.Document.Document

  constructor(options: { DocumentApp: GoogleAppsScript.Document.DocumentApp }) {
    this.DocumentApp = options.DocumentApp
    this.ui = options.DocumentApp.getUi()
    this.document = options.DocumentApp.getActiveDocument()
  }

  abstract execute(): void;

  protected moveCursor(element: GoogleAppsScript.Document.Element, offset: number): void {
    const newPosition = this.document.newPosition(element, offset)
    this.document.setCursor(newPosition)
  }

  protected getCursor(): GoogleAppsScript.Document.Position {
    return this.document.getCursor()
  }
}
