# OneDrive Document Management System using Microsoft Graph

## Register the app

Registering your web application is the first step. 

1. Sign in to the [Azure portal](https://portal.azure.com/).
2. On the top bar, click on your account and under the **Directory** list, choose the Active Directory tenant where you wish to register your application.
3. Click on **More Services** in the left hand nav, and choose **Azure Active Directory**.
4. Click on **App registrations** and choose **Add**.
5. Enter a friendly name for the application, for example 'MSGraphConnectNodejs' and select 'Web app/API' as the **Application Type**. For the Sign-on URL, enter *http://localhost:3000/login*. Click on **Create** to create the application.
6. While still in the Azure portal, choose your application, click on **Settings** and choose **Properties**.
7. Find the Application ID value and copy it to the clipboard.
8. Configure Permissions for your application:
9. In the **Settings** menu, choose the **Required permissions** section, click on **Add**, then **Select an API**, and select **Microsoft Graph**.
10. Then, click on Select Permissions and select **Sign in and read user profile** and **Send mail as a user**. Click **Select** and then **Done**.
11. In the **Settings** menu, choose the **Keys** section. Enter a description and select a duration for the key. Click **Save**.
12. **Important**: Copy the key value. You won't be able to access this value again once you leave this pane. You will use this value as your app secret.

## Configure and run the app

1. Update [`script.js/onedrive_client_id`](script.js#L11) with your app id
2. Update [`script.js/onedrive_client_secret`](script.js#L12) with your app secret
3. Update [`script.js/onedrive_refresh_token`](script.js#L13) with your application's redirect uri

Prerequisites
* [`node`](https://nodejs.org/en/) - JavaScript runtime built on Chrome V8
* [`npm`](https://docs.npmjs.com/getting-started/installing-node) - Node Package Manager

To run the app, type the following into your command line:

1. `npm install` - install application dependencies
2. `npm start` - starts the application server

## Launch the app in your browser
Once the application server has been started, open your favorite web browser to `http://localhost:3000`
