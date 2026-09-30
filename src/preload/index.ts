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
  syncStatus: call('syncStatus') as XlcodeApi['syncStatus'],
  bookStamp: call('bookStamp') as XlcodeApi['bookStamp'],
  writeConfig: call('writeConfig') as XlcodeApi['writeConfig'],
  pickFolder: call('pickFolder') as XlcodeApi['pickFolder'],
  relocateBooks: call('relocateBooks') as XlcodeApi['relocateBooks'],
  revealInFolder: call('revealInFolder') as XlcodeApi['revealInFolder'],
  readRuleFile: call('readRuleFile') as XlcodeApi['readRuleFile'],
  writeRuleFile: call('writeRuleFile') as XlcodeApi['writeRuleFile'],
  setProjectMode: call('setProjectMode') as XlcodeApi['setProjectMode'],
  openOutput: call('openOutput') as XlcodeApi['openOutput'],
};

contextBridge.exposeInMainWorld('xlcode', api);
