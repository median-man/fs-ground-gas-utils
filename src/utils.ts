export function promptUserForGithubLineURL() {
  var ui = DocumentApp.getUi()
  var response = ui.prompt('Enter the url for the GitHub source code line:')
  var result = ''
  if (response.getSelectedButton() == ui.Button.OK) {
    result = response.getResponseText().trim()
  }
  return result
}

export function promptUserForLinkText() {
  var ui = DocumentApp.getUi()
  var response = ui.prompt('Enter the text for the broken link:')
  var result = ''
  if (response.getSelectedButton() == ui.Button.OK) {
    result = response.getResponseText().trim()
  }
  return result
}

export function promptUserForBulletPoint() {
  var ui = DocumentApp.getUi()
  var button = ui.alert(
    'Add a bullet point annotation to link?',
    ui.ButtonSet.YES_NO
  )
  return button == ui.Button.YES
}

export function moveCursorToEndOfElement(element) {
  var position = DocumentApp.getActiveDocument().newPosition(
    element,
    element.getText().length
  )
  DocumentApp.getActiveDocument().setCursor(position)
}

export function promptUserForTargetPath() {
  var ui = DocumentApp.getUi()
  var response = ui.prompt('Enter the path to the link target:')
  var result = ''
  if (response.getSelectedButton() == ui.Button.OK) {
    result = response.getResponseText().trim()
  }
  return result
}
