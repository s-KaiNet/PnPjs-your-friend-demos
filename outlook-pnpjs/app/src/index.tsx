import "core-js/modules/es6.array.from.js";
import "core-js/modules/es6.array.iterator.js";
import "core-js/modules/es6.promise";
import 'es6-map/implement';
import 'es6-promise/auto';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BrowserRouter } from 'react-router-dom';
import 'whatwg-fetch';

// AppContainer is a necessary wrapper component for HMR
// import { AppContainer } from 'react-hot-loader'

/*
  Main App CSS
    - Used for introduce CSS in webpack workflow
    - In webpack Dev it will be injected as /**
    - In webpack prod it will be extracted as a separate bundled file
 */
import './../stylesheets/main.css';

/*
  Main App Container
 */
import { App } from './containers/App/App';

let isOfficeInitialized = false;

const render = (Component: any) => {
  ReactDOM.render(
    <BrowserRouter>
      <Component isOfficeInitialized={isOfficeInitialized} />
    </BrowserRouter>,
    document.getElementById('reactContainer')
  );
};

/* Render application after Office initializes */
Office.initialize = () => {
  isOfficeInitialized = true;
  render(App);
};

render(App);
