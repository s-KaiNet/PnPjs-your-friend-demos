import { BearerTokenFetchClient, FetchOptions } from "@pnp/common";
import { AadTokenProvider } from "@microsoft/sp-http";

export class GraphTokenFetchClient extends BearerTokenFetchClient {
    constructor(private tokenProvider: AadTokenProvider) {
        super(null);
    }

    public fetch(url: string, options: FetchOptions = {}): Promise<Response> {
        return this.tokenProvider.getToken(this.getResource(url))
            .then((accessToken: string) => {
                this.token = accessToken;
                return super.fetch(url, options);
            });
    }

    private getResource(url: string): string {
      const parser = document.createElement('a') as HTMLAnchorElement;
      parser.href = url;
      return `${parser.protocol}//${parser.hostname}`;
    }
}
