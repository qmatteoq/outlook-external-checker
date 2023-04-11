# Outlook add-in to check external recipients

This Outlook add-in enhances the external recipients check experience, by adding an extra layer of security to avoid leaking confidential information to people outside your organization.

## Prerequisites

You can reference [the official documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview) to identify the prerequisites for your development environment.

## Setup

1. Clone the repository
2. Open the project's folder in a terminal and run:

   ```bash
   npm install
   ```
3. In the src/launchevent/launchevent.js file, change the value of the customerDomain property to match your organizational's domain (e.g. @contoso.com).
4. Move to the Run and Debug tab in Visual Studio Code.
5. Choose the Outlook Desktop (Edge Chromium) profile.
6. Press F5 to launch Outlook and sideload the add-in.
7. Compose a new e-mail, add at least one e-mail address which doesn't belong to the domain you have specified in Step 3 and press Send.

## Deploy
This add-in is baesd on the new Office web model. If you want to deploy it in your company, you must first host it somewhere. For example, you can host it on Azure Storage [by following this guide](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-add-in-vs-code).


## Learn more
This project is described in details on the [Modern Work App Consult blog]().