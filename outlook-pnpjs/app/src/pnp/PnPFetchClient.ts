import { BearerTokenFetchClient, FetchOptions, isUrlAbsolute } from '@pnp/common';

import * as AuthenticationContext from 'adal-angular';

export class PnPFetchClient extends BearerTokenFetchClient {
  constructor(private authContext: AuthenticationContext) {
    super(null);
  }

  public fetch(url: string, options: FetchOptions = {}): Promise<Response> {
    if (!isUrlAbsolute(url)) {
      throw new Error('You must supply absolute urls to PnPFetchClient.fetch.');
    }

    return this.getToken(this.getResource(url)).then(token => {
      this.token = token;
      return super.fetch(url, options);
    });
  }

  private getToken(resource: string): Promise<string> {
    return new Promise((resolve, reject) => {
      this.authContext.acquireToken(resource, (message, token) => {
        if (!token) {
          const err = new Error(message);
          reject(err);
        } else {
          resolve(token);
        }
      });
    });
  }

  private getResource(url: string): string {
    const parser = document.createElement('a') as HTMLAnchorElement;
    parser.href = url;
    return `${parser.protocol}//${parser.hostname}`;
  }
}
