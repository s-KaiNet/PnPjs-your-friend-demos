## PnPjs with nodejs integration

### How to run
1. `npm i`
2. `npm install node-sp-auth-config -g`
3. In root directory run `sp-auth init --path ./config/private.config.json` and provide your credentials and SharePoint site url
4. In `index.ts` fix Web constructor to point to your SharePoint site url
5. `npm run start` - you will see results in a console