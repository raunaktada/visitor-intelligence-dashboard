import React from 'react';
import ReactDOM from 'react-dom/client';
import Dashboard from './Dashboard';
import './Dashboard.css';

const el = document.getElementById('root');
if (el) {
  ReactDOM.createRoot(el).render(
    <React.StrictMode>
      <Dashboard />
    </React.StrictMode>
  );
} else {
  document.body.style.background = 'red';
  document.body.textContent = 'ROOT NOT FOUND';
}
