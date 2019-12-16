interface BootstrapIssue {
  title: string
  ghUrl: string
  ghLineUrls: string[]
}

export function addBootstrapUpdateIssue() {
  const issue = createBootstrapIssueDialog()
  if (issue) {
    insertIssue(issue)
  }
}

function createBootstrapIssueDialog(): BootstrapIssue {
  const promptTitle = 'Create Bootstrap Update Issue'
  const result: BootstrapIssue = {
    title: '',
    ghUrl: '',
    ghLineUrls: []
  }
  let response: GoogleAppsScript.Base.PromptResponse
  const ui = DocumentApp.getUi()

  response = promptForGhUrl()
  if (!isResponseOk(response)) {
    return null
  }
  result.ghUrl = response.getResponseText().trim()
  result.title = lastNameOfPath(result.ghUrl)

  do {
    response = promptForGhLinePermalink()
    if (!isResponseOk(response)) {
      return
    }
    result.ghLineUrls.push(response.getResponseText().trim())
  } while (promptForAnotherPermalink() === ui.Button.YES)
  return result

  function isResponseOk(response: GoogleAppsScript.Base.PromptResponse) {
    return response.getSelectedButton() === ui.Button.OK
  }

  function lastNameOfPath(path: string) {
    return path.match(/\/([^\/]+?)$/)[1]
  }

  function promptUser(promptText) {
    return ui.prompt(promptTitle, promptText, ui.ButtonSet.OK_CANCEL)
  }

  function promptForGhUrl() {
    return promptUser('What is the GitHub URL for the activity/assignment?')
  }

  function promptForGhLinePermalink() {
    return promptUser('What is the URL for the GitHub permalink?')
  }

  function promptForAnotherPermalink() {
    const promptText = 'Add another GitHub line permalink?'
    const buttons = ui.ButtonSet.YES_NO
    return ui.alert(promptTitle, promptText, buttons)
  }
}

function insertIssue(issue: BootstrapIssue) {
  const position = DocumentApp.getActiveDocument().getCursor()
  const body = DocumentApp.getActiveDocument().getBody()  
  let index = body.getChildIndex(position.getElement())
  insertHeadingAtPosition(position)
  insertGhLink(index + 1)
  insertBlankLine(index + 2)
  insertIssueBody(index + 3)

  function insertHeadingAtPosition(position: GoogleAppsScript.Document.Position) {
    position.insertText(` - Update Bootstrap version: ${issue.title}`)
  }

  function insertGhLink(index) {
    const paragraph = body.insertParagraph(index, issue.ghUrl)
    paragraph.setLinkUrl(issue.ghUrl)
  }

  function insertBlankLine(index) {
    body.insertParagraph(index, '')
  }

  function insertIssueBody(index) {
    const paragraph = body.insertParagraph(index, '')
    const re = new RegExp(`${issue.title}/(.+)#`)
    const links = issue.ghLineUrls.map(url => {
      const lineNumber = url.match(/\d+$/)
      const fileName = url.match(re)[1]
      const text = `Line ${lineNumber}, ${fileName}`
      return { text, url }
    })
    paragraph.appendText(
      'Update the version of Bootstrap linked in the html files from 4.0 to 4.3. ('
    )
    links.forEach((link, index) => {
      if (links.length > 1 && index === links.length - 1) {
        paragraph.appendText(' and ')
      }
      paragraph.appendText(link.text)
      if (links.length > 2 && index !== links.length - 1) {
        paragraph.appendText(', ')
      }
    })
    paragraph.appendText(')')
    links.forEach(link => {
      const range = paragraph.findText(link.text)
      range
        .getElement()
        .asText()
        .setLinkUrl(
          range.getStartOffset(),
          range.getEndOffsetInclusive(),
          link.url
        )
    })
  }
}
