import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';
import App from './App';

const Office = window.Office;

Office.initialize = () => {
  ReactDOM.render(<App />, document.getElementById('root'));
};

