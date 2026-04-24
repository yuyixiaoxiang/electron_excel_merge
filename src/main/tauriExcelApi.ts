import { Menu, MenuItem, Submenu } from '@tauri-apps/api/menu';
import type { CellChange, SaveChangesRequest, WorkspaceTabMenuEvent } from './preload';
import {
  computeThreeWayDiff,
  debugLog,
  getCliThreeWayInfo,
  getDebugLogPathValue,
  getSheetData,
  getThreeWayRow,
  getThreeWayRows,
  listExcelFilesInFolder,
  loadWorkbook,
  openFile,
  openThreeWay,
  pickFolder,
  saveChanges,
  saveMergeResult,
} from './excelBackend';

const workspaceTabListeners = new Set<(payload: WorkspaceTabMenuEvent) => void>();

export const emitWorkspaceNewTab = (payload: WorkspaceTabMenuEvent) => {
  workspaceTabListeners.forEach((listener) => listener(payload));
};

const initAppMenu = async () => {
  try {
    const fileMenu = await Submenu.new({
      id: 'workspace-file-menu',
      text: 'File',
      items: [
        await MenuItem.new({
          id: 'workspace-folder',
          text: 'Excel 文件夹比较',
          action: () => emitWorkspaceNewTab({ kind: 'folder' }),
        }),
        await MenuItem.new({
          id: 'workspace-diff',
          text: 'Excel 比较',
          action: () => emitWorkspaceNewTab({ kind: 'diff' }),
        }),
        await MenuItem.new({
          id: 'workspace-merge',
          text: 'Merge 模式',
          action: () => emitWorkspaceNewTab({ kind: 'merge' }),
        }),
      ],
    });

    const menu = await Menu.new({
      id: 'emerge-app-menu',
      items: [fileMenu],
    });

    await menu.setAsAppMenu();
  } catch (error) {
    console.warn('failed to initialize app menu', error);
  }
};

window.excelAPI = {
  pickFolder,
  listExcelFilesInFolder,
  openFile,
  loadWorkbook,
  saveChanges: async (req: SaveChangesRequest | CellChange[]): Promise<void> => {
    await saveChanges(req);
  },
  openThreeWay,
  getSheetData,
  computeThreeWayDiff,
  saveMergeResult,
  getCliThreeWayInfo,
  getThreeWayRow,
  getThreeWayRows,
  debugLog,
  getDebugLogPath: getDebugLogPathValue,
  onWorkspaceNewTab: (handler: (payload: WorkspaceTabMenuEvent) => void): (() => void) => {
    workspaceTabListeners.add(handler);
    return () => {
      workspaceTabListeners.delete(handler);
    };
  },
};

void initAppMenu();
