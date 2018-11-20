## Outlook add-in with PnPjs and adal.js

## How to run

1. Create a new app registation in Azure AD. Enable implict flow, add `https://localhost:3000/` as a valid redirect url
2. Add API permissons to the app: SharePoint: `AllSites.FullControl`
3. Under `./webpack/dev.env.js` change `SP_SITE_URL` to point to your SharePoint site
4. Under `./app/src/adal/adalConfig.ts` replace values with yours tenant id and client id from step 1
5. Copy full path to file `./outlook-pn-p-manifest.xml`. Open Outlook, click on "Get Add-ins" in the ribbon. Then My Add-ins -> Custom Add-ins -> Add custom add-in from file. Select path to `./outlook-pn-p-manifest.xml`. Restart Outlook.
6. `npm i`
7. `npm run start` - starts a web server with your add-in. In Outlook create a new mail, in the ribbon a new button "Insert template" will appear. Click on the button, login and add templates into the message body.