# Legacy Normalizer

This is a library to normalize legacy data.

## Reminders

Open your terminal and check if you have node.js installed in your machine

```
$ node --version
```

If not, you can [download it here](https://nodejs.org/). Open it and follow the steps.
After finished, run the same command again

Don't forget to download the necessary dependencies

```
$ npm install
```

## Steps

1. Open VS Code
2. Open the terminal (find "Terminal" at the top of your screen then click "New Terminal")
3. In the terminal, type `npm install` then enter
4. Add the spreadsheet you want to normalize to the spreadsheets folders on this directory.
   Your file must follow the template `spreadsheets/layout-surety-v2.xlsx`
5. Inside the `src` folder, find the script you want to use. Open it then write the name of
   your file in the variable `fileName`, e.g., `const fileName = 'austral-legacy'`
6. Go to the terminal then run `npm run <script-name>`
7. Congratulations! Now you have the normalized file inside the spreadsheets folder
