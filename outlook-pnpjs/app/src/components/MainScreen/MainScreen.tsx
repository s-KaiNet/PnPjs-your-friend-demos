import { Pivot, PivotItem, PivotLinkSize } from 'office-ui-fabric-react/lib/Pivot';
import * as React from 'react';

import { Header } from '../../components/Header/Header';
import { Documents } from '../../containers/Documents/Documents';

export const MainScreen: React.StatelessComponent<{}> = () => (
  <div>
    <Header />
    <Documents />
  </div>
);
