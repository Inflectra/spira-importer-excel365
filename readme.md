## Setting up dev Environment

The source code of the add-in is located in our internal GIT at Importers\MS-Office365

- **1** -- Refer to the page https://docs.microsoft.com/en-us/office/dev/add-ins/overview/learning-path-beginner
- **2** -- In step 2 choose node JS and Visual Studio Code
-- **2.1** -- If you don't have Node.js and npm installed, follow the instructions from and install them https://docs.microsoft.com/en-us/office/dev/add-ins/overview/set-up-your-dev-environment
-- **2.2** -- Run the cmd command: npm install -g yo generator-office
- **3** -- Open the root folder in VS Code, click Ctrl + ` to open the terminal - inside VS Code, and run the command "npm install"
- **4** -- Run the command npm start to run your project.

## Debugging

In Excel on Desktop:
* Locate a small arrow at top right of the add-in (<). Click on it and select 'Attach Debugger'. This will open a console browser-style windows, where you can check the add-in running.

## Troubleshooting

*Error: Sorry, we can't load the add-in. Please make sure you have network and/or Internet connectivity.*: That means your dev server is not started. To start it, open a new Command Prompt, point it to your folder, and run the command **npm run dev-server**. Keep the console opened while you are working.