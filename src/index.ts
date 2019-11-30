import { insertBrokenLink } from './insert-broken-link'
import { insertLinkToIssue } from './insert-link-to-issue'
import { addIssueHeadings } from './add-issue-headings'

declare var global: any
global.onOpen = onOpen
global.insertBrokenLink = insertBrokenLink
global.insertLinkToIssue = insertLinkToIssue
global.addIssueHeadings = addIssueHeadings

function onOpen() {
  createUtilitiesMenu()
}

function createUtilitiesMenu() {
  DocumentApp.getUi()
    .createMenu('Utilities')
    .addItem('Insert Broken Link Issue', 'insertBrokenLink')
    .addItem('Insert Link To Issue', 'insertLinkToIssue')
    .addItem('Add issue headings', 'addIssueHeadings')
    // .addItem('Bootstrap update issue', 'insertBootstrapIssue') // TODO: finish this feature
    //    .addItem('Test new form', 'newBrokenLinkDialog')
    .addToUi()
}
