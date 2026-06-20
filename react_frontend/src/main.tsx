import React from 'react';
import { createRoot } from 'react-dom/client';
import './styles.css';
import { VietmaxApp } from './vietmax/VietmaxApp';

createRoot(document.getElementById('root') as HTMLElement).render(
  <React.StrictMode>
    <VietmaxApp />
  </React.StrictMode>,
);
