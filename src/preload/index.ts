import { contextBridge, ipcRenderer } from 'electron';
import type { XlcodeApi } from '../shared/api';

const call =
  (channel: string) =>
  (...args: unknown[]) =>
    ipcRenderer.invoke(channel, ...args);

const api: XlcodeApi = {
  platform: process.platform,
  pickProject: call('pickProject') as XlcodeApi['pickProject'],
  loadProject: call('loadProject') as XlcodeApi['loadProject'],
  initProject: call('initProject') as XlcodeApi['initProject'],
  refreshTree: call('refreshTree') as XlcodeApi['refreshTree'],
  createBook: call('createBook') as XlcodeApi['createBook'],
  build: call('build') as XlcodeApi['build'],
  sync: call('sync') as XlcodeApi['sync'],
  openInExcel: call('openInExcel') as XlcodeApi['openInExcel'],
  openTerminal: call('openTerminal') as XlcodeApi['openTerminal'],
  bookLocks: call('bookLocks') as XlcodeApi['bookLocks'],
  readConfig: call('readConfig') as XlcodeApi['readConfig'],
  writeConfig: call('writeConfig') as XlcodeApi['writeConfig'],
  revealInFolder: call('revealInFolder') as XlcodeApi['revealInFolder'],
  readRuleFile: call('readRuleFile') as XlcodeApi['readRuleFile'],
  writeRuleFile: call('writeRuleFile') as XlcodeApi['writeRuleFile'],
};

contextBridge.exposeInMainWorld('xlcode', api);
