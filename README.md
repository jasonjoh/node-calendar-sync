# Node.js Sample for Calendar Sync with Office 365 #

## Installation ##

    npm install
    
## Running the sample

1. Register a new app at https://apps.dev.microsoft.com
  1. Copy the **Application Id** and paste this value for the `clientId` value in `authHelper.js`.
  1. Click the **Generate New Password** button and copy the password. Paste this value for the `clientSecret` value in `authHelper.js`.
  1. Click the **Add Platform** button and choose **Web**. Enter `http://localhost:3000/authorize` for the **Redirect URI**.
  1. Click **Save**.
1. Save changes to `authHelper.js` and start the app:

        npm start
    
1. Open your browser and browse to http://localhost:3000.

## Copyright ##

Copyright (c) Microsoft. All rights reserved.

----------
Connect with me on Twitter [@JasonJohMSFT](https://twitter.com/JasonJohMSFT)

Follow the [Outlook Dev Blog](http://blogs.msdn.com/b/exchangedev/)