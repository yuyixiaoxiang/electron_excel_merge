/// <reference path="./global.d.ts" />
/**
 * React 入口：在 #root 容器中挂载 App 组件。
 *
 * global.d.ts 中声明的 window.excelAPI 会在整个 React 应用中可用。
 */
import React from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';

const container = document.getElementById('root');
if (container) {
  const root = createRoot(container);
  root.render(<App />);
}
