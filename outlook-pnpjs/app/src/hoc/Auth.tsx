import { graph } from '@pnp/graph';
import { sp } from '@pnp/sp';
import * as AuthenticationContext from 'adal-angular';
import * as React from 'react';

import { adalConfig } from '../adal/adalConfig';
import { LoginError } from '../components/LoginError/LoginError';
import { LoginInProgress } from '../components/LoginInProgress/LoginInProgress';
import { PnPFetchClient } from '../pnp/PnPFetchClient';

export const authContext: AuthenticationContext = new AuthenticationContext(adalConfig);

interface IState {
  authenticated: boolean;
  renewIframe: boolean;
  errorMessage: string;
  hasError: boolean;
}

export function withAuth<TOriginalProps>(
  WrappedComponent: React.ComponentClass<TOriginalProps> | React.StatelessComponent<TOriginalProps>
): React.ComponentClass<TOriginalProps> {
  return class Auth extends React.Component<TOriginalProps, IState> {
    constructor(props: TOriginalProps) {
      super(props);

      this.state = {
        authenticated: false,
        renewIframe: false,
        hasError: false,
        errorMessage: null
      };
    }

    public componentWillMount(): void {
      authContext.handleWindowCallback();

      if (authContext.isCallback(window.location.hash)) {
        this.setState({
          renewIframe: true
        });
        return;
      }

      if (!authContext.getCachedToken(adalConfig.clientId) || !authContext.getCachedUser()) {
        authContext.login();
      } else if (authContext.getLoginError()) {
        this.setState({
          hasError: true,
          errorMessage: authContext.getLoginError()
        });
      } else {
        this.setState({
          authenticated: true
        });

        sp.setup({
          sp: {
            fetchClientFactory: () => new PnPFetchClient(authContext),
            baseUrl: process.env.SP_SITE_URL
          }
        });

        graph.setup({
          graph: {
            fetchClientFactory: () => {
              return new PnPFetchClient(authContext);
            },
          }
        });
      }
    }

    public render(): JSX.Element {
      if (this.state.renewIframe) {
        return <div>hidden renew iframe - not visible</div>;
      }

      if (this.state.authenticated) {
        return <WrappedComponent {...this.props} />;
      }

      if (this.state.hasError) {
        return <LoginError message={this.state.errorMessage} />;
      }

      return <LoginInProgress />;
    }
  };
}
