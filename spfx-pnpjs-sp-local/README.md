## SharePoint Framework with PnPjs and local SharePoint development

### How to run

1. `npm i`
2. `npm install node-sp-auth-config -g`
3. In root directory run `sp-auth init --path ./config/private.json` and provide your credentials and SharePoint site url
4. `npm run start` - starts proxy and spfx serve
5. Open your SharePoint site and add Clients list with text fields `Address`, `Email`, `Company`. Enter test data. Alternatively modify `ClientList.tsx` to point to your list. 
6. Add web part in local workbench and see the results.