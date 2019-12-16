import { insertBrokenLink } from './insert-broken-link'
import { insertLinkToIssue } from './insert-link-to-issue'
import { addIssueHeadings } from './add-issue-headings'
import { addBootstrapUpdateIssue } from './add-bootstrap-issue'
import { insertActivityLink } from './insert-activity-link'
import { createES6ActivityIssue } from './issue-es6-activity'

declare var global: any
global.onOpen = onOpen
global.insertBrokenLink = insertBrokenLink
global.insertLinkToIssue = insertLinkToIssue
global.addIssueHeadings = addIssueHeadings
global.addBootstrapUpdateIssue = addBootstrapUpdateIssue
global.insertActivityLink = insertActivityLink
global.createES6ActivityIssue = createES6ActivityIssue

function onOpen() {
  createUtilitiesMenu()
}

interface UtilitiesMenuItem {
  text: string
  onClick: string
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
    { text: 'Insert Activity Link', onClick: 'insertActivityLink' },
    { text: 'ES6 Activity Refactor Issue', onClick: 'createES6ActivityIssue' }
  ]
  const menu = DocumentApp.getUi().createMenu('Utilities')
  utilitiesMenuItems.forEach(item => menu.addItem(item.text, item.onClick))
  menu.addToUi()
}
