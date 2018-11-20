import { bootstrap } from 'pnp-auth';
import { sp, Web } from '@pnp/sp';
import { SPFetchClient } from "@pnp/nodejs";

bootstrap(sp, './config/private.config.json');

let web = new Web('https://mastaq.sharepoint.com/sites/pnpjs');

web.get().then((data) => {
    console.log(data);
})

/*
sp.setup({
    sp: {
        fetchClientFactory : () => new SPFetchClient('https://sp.url', 'client id', 'client secret')
    }
});
*/