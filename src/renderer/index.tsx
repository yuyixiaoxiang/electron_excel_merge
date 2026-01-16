/// <reference path="./global.d.ts" />
/**
 * React 入口：在 #root 容器中挂载 App 组件。
 *
 * global.d.ts 中声明的 window.excelAPI 会在整个 React 应用中可用。
 */
import React from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';

// 全局样式：让 html / body / #root 占满窗口，并禁用最外层滚动条，
// 只保留表格内部的滚动条（ExcelTable 自己的 scroll）。
const styleTag = document.createElement('style');
styleTag.textContent = `
  html, body, #root {
    margin: 0;
    padding: 0;
    height: 100%;
    overflow: hidden;
  }
`;
document.head.appendChild(styleTag);

const container = document.getElementById('root');
if (container) {
  const root = createRoot(container);
  root.render(<App />);
}
