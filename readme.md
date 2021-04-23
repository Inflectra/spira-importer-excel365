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

* Close everything, stop the dev server
* Re-open your files. Run the add-in. Locate a small arrow at top right of the add-in (<). Click on it and select 'Attach Debugger'. This will open a console browser-style windows, where you can check the add-in running.
* If that not works, you can try to install the Microsoft Edge Dev Tools

## Troubleshooting

*Error: Sorry, we can't load the add-in. Please make sure you have network and/or Internet connectivity.*: That means your dev server is not started. To start it, open a new Command Prompt, point it to your folder, and run the command **npm run dev-server**. Keep the console opened while you are working.