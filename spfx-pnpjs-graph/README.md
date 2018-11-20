## SharePoint Framework with PnPjs and MS Graph queries

### How to run
1. `npm i`
2. `gulp bundle --ship` then `gulp package-solution --ship`
3. Upload resulting app file into the app catalog
4. Approve app permission request from modern SharePoint Admin (API Permissions)
5. `npm run start` - runs spfx gulp serve
6. Open hosted workbench `<sp site>/_layouts/15/workbench.aspx`, add web part and see the results