## Setting up dev Environment

- **1** -- Refer to the page https://docs.microsoft.com/en-us/office/dev/add-ins/overview/learning-path-beginner
- **2** -- In step 2 choose node JS and Visual Studio Code
-- **2.1** -- If you don't have Node.js and npm installed, follow the instructions from and install them https://docs.microsoft.com/en-us/office/dev/add-ins/overview/set-up-your-dev-environment
-- **2.2** -- Run the cmd command: npm install -g yo generator-office
- **3** -- Open the root folder in VS Code, click Ctrl + ` to open the terminal - inside VS Code, and run the command "npm install"
- **4** -- Run the command npm start to run your project.

## Debugging

In Excel on Desktop:
* Follow the instructions of this article: https://docs.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10, i.e: 

Run the command Get-AppxPackage Microsoft.Win32WebViewHost
Copy the PackageFullName value
Run the command setx JS_DEBUG <PackageFullName> with the full name you just got

* If that not works, you can try to install the Microsoft Edge Dev Tools
* To DEBUG in Visual Studio Code, install and follow the instructions of this extension: https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger

## Troubleshooting

*Error: Sorry, we can't load the add-in. Please make sure you have network and/or Internet connectivity.*: That means your dev server is not started. To start it, open a new Command Prompt, point it to your folder, and run the command **npm run dev-server**. Keep the console opened while you are working.

## Deploying
1. Validate - run `npm run validate`
2. Build - run `npm run build`. This will generate files in the \dist folder. These, and the manifest.xml, are the actual distributed files. Nothing else.

## Testing the addin - for testers
Overall, we will be following the steps for [windows](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) and the [web](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

1. Run the validate and build steps above
2. Make a copy of the manifest.xml file and change active urls (either pointing to localhost or files.inflectra.com) so that the base URL is `https://files.inflectra.com/office365/excel-staging/`
3. Find a shared network location to place the manifest XML and paste the file there
4. Generate the TrustNetworkShareCatalog.reg file as described in the windows guide above. Run if on a Windows machine, then reboot.
5. then task a web admin with replacing the files in `https://files.inflectra.com/office365/excel-staging/` with all the files generated at build in the `dist` folder
6. follow the steps to load the addin into Excel via Windows or the web