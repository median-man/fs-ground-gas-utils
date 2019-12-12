import { insertBrokenLink } from './insert-broken-link'
import { insertLinkToIssue } from './insert-link-to-issue'
import { addIssueHeadings } from './add-issue-headings'
import { addBootstrapUpdateIssue } from './add-bootstrap-issue'
import { insertActivityLink } from './insert-activity-link'

declare var global: any
global.onOpen = onOpen
global.insertBrokenLink = insertBrokenLink
global.insertLinkToIssue = insertLinkToIssue
global.addIssueHeadings = addIssueHeadings
global.addBootstrapUpdateIssue = addBootstrapUpdateIssue
global.insertActivityLink = insertActivityLink

function onOpen() {
  createUtilitiesMenu()
}

interface UtilitiesMenuItem {
  text: string;
  onClick: string;
}

function createUtilitiesMenu() {
  // onClick string must be key used when assigning the function to the global
  // object
  const utilitiesMenuItems: UtilitiesMenuItem[] = [
    { text: 'Insert Broken Link Issue', onClick: 'insertBrokenLink' },
    { text: 'Insert Link To Issue', onClick: 'insertLinkToIssue' },
    { text: 'Add issue headings', onClick: 'addIssueHeadings' },
    { text: 'Bootstrap Update Issue', onClick: 'addBootstrapUpdateIssue' },
    { text: 'Bootstrap Update Issue', onClick: 'addBootstrapUpdateIssue' },
    { text: 'Insert Activity Link', onClick: 'insertActivityLink' }
  ]
  const menu = DocumentApp.getUi().createMenu('Utilities')
  utilitiesMenuItems.forEach(item => menu.addItem(item.text, item.onClick))
  menu.addToUi()
}
