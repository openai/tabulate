# OpenAI API Excel integration

This repository contains an example OpenAI API integration for Excel. It allows users to query the API to automatically generate Excel tables about topics.

See a [demo video](https://beta.openai.com/?app=productivity&example=4_4_0)

The integration is an Excel TaskPane Add-in, which is structured as an HTML / CSS / Javascript web app running in an iframe. See the following links for more info:
- https://docs.microsoft.com/en-us/office/dev/add-ins/overview/learning-path-beginner
- https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-core-concepts

## Setup

Add your OpenAI API key and organization at the top of `excel-addin/src/taskpane.js` (search for `***KEY HERE***` and `***ORG HERE***`)

To start the local development server from the `excel-addin` directory:
- `brew install node@12` (Node LTS)
- `npm install`
- `npm run dev-server`

Open Excel for the web. Click "Insert" Menu (Ribbon) > Click "Office Add-ins" > Click "Upload My Add-in" in the upper right corner > Select `excel-addin/manifest.xml` ([source](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#sideload-an-office-add-in-in-office-on-the-web))

You should see a new "OpenAI API" command group on the "Home" ribbon; click the "Tabulate" button to open the sidebar with API commands
