import * as React from 'react';
import * as ReactDOM from 'react-dom';
import './styles.css';
import { AppFunc } from './app-func';

// 1. Functional Component sample
ReactDOM.render(<AppFunc />, document.getElementById('app'));

// 2. Class Component sample
// ReactDOM.render(<App />, document.getElementById('app'));