import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import './styles/themes.css';
import './styles/app.css';
import './shell.css';
// 先设置编辑器宿主，再加载界面
import './host';
import { App } from './App';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
