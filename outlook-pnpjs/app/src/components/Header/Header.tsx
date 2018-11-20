import * as React from 'react';

import { authContext } from '../../hoc/Auth';
import * as styles from './Header.css';

export const Header: React.StatelessComponent<{}> = () => {
  const user = authContext.getCachedUser().profile.name;

  return (
    <div className={styles.header}>
      Hello, <i>{user}</i>!
    </div>
  );
};
