# 🚀 Welcome to your new awesome project!

This project has been created using **webpack scaffold**, you can now run

```
npm run build
```

or

```
yarn build
```

to bundle your application

## Deploy Google Apps Script

This project is built using
[clasp](https://developers.google.com/apps-script/guides/clasp).

The script is intended to bound to a Google Doc. Set the `scriptId` field in
[.clasp.json](./.clasp.json) and run `clasp push` after running `npm run build`
to deploy and start using the utility. (May requiring installing and logging in
with `clasp` first.)
