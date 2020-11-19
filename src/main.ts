import { EventEmitter } from 'events';
import util from 'util';
import fs from 'fs';

// @ts-ignore
import * as winax from 'winax';
import { CrebassObject } from './interface';

function proxyCOM(target: any) {
  const getter: Record<string, boolean> = {};

  if (target === undefined) return undefined;
  for (let key of [...target.__type])
    if (
      (key.invkind === 2) && (target[key.name] != null) &&
      (target[key.name].__value !== '[object]')
    )
      getter[key.name] = true;

  return new Proxy(target, {
    get: function (target, name) {
      target = target[name];
      return (name in getter) ? target.valueOf() : target;
    },
    set: function (target, name, value) {
      target[name] = value;
      return true;
    }
  });
}

function waitFor(filter: Function, notnull: boolean = false, timeout: number = 30000) {
  return new Promise((resolve, reject) => {
    var start = Date.now();

    setTimeout(async function check() {
      if (timeout && ((Date.now() - start) >= timeout))
        return filter ? reject(`Timeout - ${timeout}ms`) : resolve();

      if (filter) {
        try {
          var result = await filter();
        } catch (error) { return reject(error); }

        if (notnull) {
          if (result != null) return resolve(result);
        } else if (result) return resolve(result);
      }

      setTimeout(check);
    });
  });
}

enum FrameName {
  _blank = 'blank',
  _parent = '_parent',
  _self = '_self',
  _top = '_top'
}

enum RefreshConstants {
  REFRESH_NORMAL,
  REFRESH_IFEXPIRED,
  REFRESH_COMPLETELY
}

interface IEContext {
  // ExecWB: (cmdId: string, cmdexecOpt: string, pvaln: Buffer, pvaOut: Buffer);
  AddressBar: boolean;
  Application: string;
  Busy: boolean;
  Container: object | null;
  Document?: object;
  FullName: string;
  FullScreen: boolean;
  Height: string;
  Left: string;
  Locationname: string;
  LocationURL: string;
  MenuBar: boolean;
  Offline: boolean;
  Parent: string;
  Path: string;
  ReadyState: number;
  RegisterAsBrowser: boolean;
  RegisterAsDropTarget: boolean;
  Resizable: boolean;
  Silent: boolean;
  StatusBar: boolean;
  TheaterMode: boolean;
  ToolBar: number;
  Top: number;
  TopLevelContainer: boolean;
  Type?: object;
  Visible: boolean;
  Width: number;

  GoBack(): void;
  GoForward(): void;
  GoHome(): void;
  GoSearch(): void;
  Navigate(url: string, flags?: string, targetFrameName?: FrameName | string, postData?: string, headers?: string): void;
  Navigate2(url: string, flags?: string, targetFrameName?: FrameName | string, postData?: string, headers?: string): void;
  Refresh(level?: RefreshConstants): void;
  Refresh2(level?: RefreshConstants): void;
  Stop(): void;
  Quit(): void;
}

interface IEWindow {
  alert(message?: any): void;
  blur(): void;
  clearInterval(handle?: number): void;
  clearTimeout(handle?: number): void;
  close(): void;
  confirm(message?: string): boolean;
  execScript(code: string, language?: string): void;
  focus(): void;
  moveBy(x: number, y: number): void;
  moveTo(x: number, y: number): void;
  name: string;
  fafa: string;
  navigate(url: string): void;
  open(url?: string, name?: FrameName | string, features?: string, replace?: boolean): void;
  // opener
  // parent
  prompt(message?: string, _default?: string): string | null;
  resizeBy(x: number, y: number): void;
  resizeTo(width: number, height: number): void;
  scroll(options?: ScrollToOptions): void;
  scroll(x: number, y: number): void;
  scrollBy(options?: ScrollToOptions): void;
  scrollBy(x: number, y: number): void;
  scrollTo(options?: ScrollToOptions): void;
  scrollTo(x: number, y: number): void;
  // self
  setInterval(handler: TimerHandler, timeout?: number, ...args: any[]): number;
  setTimeout(handler: TimerHandler, timeout?: number, ...args: any[]): number;
  // showHelp
  // showModalDialog
  // status
}

class CrebassPromise {
  constructor(code: string, IE: InternetExplorer) {
    const id = Date.now();
    const promise = new Promise((resolve, reject) => {
      crebassPromise[id] = {
        resolve,
        reject
      };

      // code = code.replace(/;("?)$/, '$1');
      // IE.window.execScript('document.asdf=1');
      // IE.document.asdf = code;
      IE.document.__crebass__promise(id, code);
      // console.log(IE.document.asdf);
      // (IE.window as any).execScript(`document.__crebass__.resolve(${id}, (function(){ try { return Function("${code}");  debugger;} catch(err) { return JSON.stringify(err) } })());`)
    });
    return promise;
  }
}

class InternetExplorer extends EventEmitter implements IEContext {
  _context: ActiveXObject & IEContext;
  constructor() {
    (super() as any)._target = this._context = new (ActiveXObject as any)('InternetExplorer.Application', {
      activate: true,
      async: true
    });
  }

  /** Helper Methods **/
  async eval(code: string) {
    return new CrebassPromise(code, this);
  }

  /** Helper Objects **/
  get document() {
    return proxyCOM(this._context.Document);
  }

  get window(): IEWindow {
    return proxyCOM(this.document.defaultView);
  }

  /** IEObject Fields **/

  get AddressBar() {
    return this._context.AddressBar;
  }
  set AddressBar(val) {
    this._context.AddressBar = val;
  }
  get Application() {
    return this._context.Application;
  }
  set Application(val) {
    this._context.Application = val;
  }
  get Busy() {
    return this._context.Busy;
  }
  set Busy(val) {
    this._context.Busy = val;
  }
  get Container() {
    return this._context.Container;
  }
  set Container(val) {
    this._context.Container = val;
  }
  get Document() {
    return this._context.Document;
  }
  set Document(val) {
    this._context.Document = val;
  }
  get FullName() {
    return this._context.FullName;
  }
  set FullName(val) {
    this._context.FullName = val;
  }
  get FullScreen() {
    return this._context.FullScreen;
  }
  set FullScreen(val) {
    this._context.FullScreen = val;
  }
  get Height() {
    return this._context.Height;
  }
  set Height(val) {
    this._context.Height = val;
  }
  get Left() {
    return this._context.Left;
  }
  set Left(val) {
    this._context.Left = val;
  }
  get Locationname() {
    return this._context.Locationname;
  }
  set Locationname(val) {
    this._context.Locationname = val;
  }
  get LocationURL() {
    return this._context.LocationURL;
  }
  set LocationURL(val) {
    this._context.LocationURL = val;
  }
  get MenuBar() {
    return this._context.MenuBar;
  }
  set MenuBar(val) {
    this._context.MenuBar = val;
  }
  get Offline() {
    return this._context.Offline;
  }
  set Offline(val) {
    this._context.Offline = val;
  }
  get Parent() {
    return this._context.Parent;
  }
  set Parent(val) {
    this._context.Parent = val;
  }
  get Path() {
    return this._context.Path;
  }
  set Path(val) {
    this._context.Path = val;
  }
  get ReadyState() {
    return this._context.ReadyState;
  }
  set ReadyState(val) {
    this._context.ReadyState = val;
  }
  get RegisterAsBrowser() {
    return this._context.RegisterAsBrowser;
  }
  set RegisterAsBrowser(val) {
    this._context.RegisterAsBrowser = val;
  }
  get RegisterAsDropTarget() {
    return this._context.RegisterAsDropTarget;
  }
  set RegisterAsDropTarget(val) {
    this._context.RegisterAsDropTarget = val;
  }
  get Resizable() {
    return this._context.Resizable;
  }
  set Resizable(val) {
    this._context.Resizable = val;
  }
  get Silent() {
    return this._context.Silent;
  }
  set Silent(val) {
    this._context.Silent = val;
  }
  get StatusBar() {
    return this._context.StatusBar;
  }
  set StatusBar(val) {
    this._context.StatusBar = val;
  }
  get TheaterMode() {
    return this._context.TheaterMode;
  }
  set TheaterMode(val) {
    this._context.TheaterMode = val;
  }
  get ToolBar() {
    return this._context.ToolBar;
  }
  set ToolBar(val) {
    this._context.ToolBar = val;
  }
  get Top() {
    return this._context.Top;
  }
  set Top(val) {
    this._context.Top = val;
  }
  get TopLevelContainer() {
    return this._context.TopLevelContainer;
  }
  set TopLevelContainer(val) {
    this._context.TopLevelContainer = val;
  }
  get Type() {
    return this._context.Type;
  }
  set Type(val) {
    this._context.Type = val;
  }
  get Visible() {
    return this._context.Visible;
  }
  set Visible(val) {
    this._context.Visible = val;
  }
  get Width() {
    return this._context.Width;
  }
  set Width(val) {
    this._context.Width = val;
  }

  GoBack() {
    return this._context.GoBack();
  }
  GoForward() {
    return this._context.GoForward();
  }
  GoHome() {
    return this._context.GoHome();
  }
  GoSearch() {
    return this._context.GoSearch();
  }
  Navigate(url: string, flags?: string | undefined, targetFrameName?: string | undefined, postData?: string | undefined, headers?: string | undefined) {
    return this._context.Navigate(url, flags, targetFrameName, postData, headers);
  }
  Navigate2(url: string, flags?: string | undefined, targetFrameName?: string | undefined, postData?: string | undefined, headers?: string | undefined) {
    return this._context.Navigate2(url, flags, targetFrameName, postData, headers);
  }
  Refresh(level?: RefreshConstants | undefined) {
    return this._context.Refresh(level);
  }
  Refresh2(level?: RefreshConstants | undefined) {
    return this._context.Refresh2(level);
  }
  Stop() {
    this._context.Stop();
  }
  async Quit() {
    this._context.Quit();
    winax.release(this._context);

    this.emit('close');
  }
}

async function sleep(ms: number) {
  return new Promise(res => {
    setTimeout(() => res(), ms);
  });
}

function attachBrowser(IE: InternetExplorer) {
  for (let event of ['uncaughtException', 'unhandledRejection', 'SIGINT', 'exit'])
    process.on(event, IE.Quit);
}

const crebassPromise: Record<number, {
  resolve: Function, reject: Function
}> = {};

const crebassObject: CrebassObject = {
  log: (track: string, ...content: any[]) => {
    if (track === '[object]') {
      track = proxyCOM(track);
    }
    console.log(`[CREBASS] ${track}: ${content}`);
  },
  logJson: (track: string, json: string) => {
    console.log('logJson', json);
    try {
      const data = JSON.parse(json);
      if (!data.length) return;
      console.log.apply([`CREBASS`, track, data]);
    } catch(e) {
      console.error(e);
    };
  },
  onerror: (message, URL, line, column, error) => {
    console.error('onerror', message, URL, line, column, error);
  },
  testMethod: (a: any, b: any, c: any, d: any, e: any) => {
    console.log('testmethod success');
  },
  ping: () => {
    return Date.now();
  },
  resolve: (id, result) => {
    const promise = crebassPromise[id];
    const data = JSON.parse(result);
    const func = data[0] === 0 ? promise.resolve : promise.reject;
    func(data[1]);
  }
};

const crebassScript = `(function(document, window) {
  if (document.__crebass__ && !document.__crebass__init) {
    console.log('prepare to crebass install');
    document.__crebass__.log('INSTALL', 'prepare to crebass install');
    document.__crebass__init = true;

    let log = console.log;
    let warn = console.warn;
    let error = console.error;

    let crebassLog = function() {
      document.__crebass__.logJson.call(this, 'console.log', JSON.stringify(arguments));
      log.apply(this, arguments);
    };

    document.__crebass__promise = function(id, code) { document.__crebass__.resolve(id, (function(){ try { return JSON.stringify(Array(0,eval(code))); } catch(err) { return JSON.stringify(Array(1,JSON.stringify(err))) } })()) };

    window.onerror = function (message, URL, line, column, error) {
      console.trace(error);
      console.error(error);
      if (! (error instanceof Error)) {
          error = new Error( message );
          error.fileName = URL,  error.lineNumber = line,  error.columnNumber = column;
      }
      document.__crebass__.onerror(message, URL, line, column, JSON.stringify(error));
  };

    function watchConsoleObject() {
      if (window.console.log !== crebassLog) {
        window.console.log = crebassLog;
      }
      setTimeout(watchConsoleObject, 1000);
    }
    watchConsoleObject();
  }
})(document, window);
`;

(async() => {
  const IE = new InternetExplorer();
  IE.Visible = true;
  IE.MenuBar = false;
  IE.ToolBar = 0;
  var connectionPoints = winax.getConnectionPoints(IE._context);
  var connectionPoint = connectionPoints[0];
  connectionPoint.advise({
    BeforeNavigate2: async function (_1: unknown, referrer: string | undefined, _3: string, frameName: string, unknownFlags: number, url: any, g: any, h: any) {
      console.log('before navigate', frameName, referrer, unknownFlags, url, g, h);
      if (unknownFlags !== 0) {
        // pending.push(url)
        // await waitForDocument();
        // await sleep(1000);
        // IE.window.execScript(`document.testMethod = null;`);

        // await sleep(1000);
        // IE.document.testMethod = (a: any, b: any, c: any, d: any, e: any) => {
        //   console.log('testMethod test', a, b, c, d, e);
        // }
      }
    },
    NavigateComplete2: async function (a: any, b: any, c: any) {
      console.log('NavigateComplete2', a, b, c);
      IE.window.execScript(`document.__crebass__ = null;document.__crebass__promise = null;`);
      await sleep(200);
      IE.document.__crebass__ = new Proxy(crebassObject, {});
      IE.window.execScript(crebassScript);
      const retaku_track = fs.readFileSync('./src/jquery.min.js', 'utf8');
      console.log('#1 jquery.min.js res', await IE.eval(retaku_track));
      try {
        const retaku_k = fs.readFileSync('./src/lodash.min.js', 'utf8');
        console.log('#2-lodash.min.js res', await IE.eval(retaku_k));
      } catch (e) {
        console.error('#2-lodash.min.js error', e);
      }
    },
    DocumentComplete : async function (a: any, b: any, c: any) {
      console.log('DocumentComplete', a, b, c);
    }

  })
  await sleep(1000);
  IE.Navigate2('http://87mm.co.kr/');
  await sleep(3000);

  // IE.window.execScript(`.testMethod='12345'; document.ping = null;`);

  // await sleep(1000);
  // (IE.window as any).testMethod = (x: string) => {
  //   console.log('test method', x);
  // };

  // (IE.window as any).ping = () => {
  // };

  (function keepalive() {
    setInterval(() => {
      try {
        (IE.window as any)
      } catch {}
    }, 10);
  })();


})();
  // (async () => {
    // attachBrowser(IE);
  // await sleep(1000);
  // IE.window.execScript(`document.testMethod='12345'`);
  // await sleep(1000);
  // console.log(IE.document.window);
  // (IE as any)._context.PutProperty('window.fafa', 123456);
  // try {
  //   while(1) {
  //     // IE.document.testMethod = () => {
  //     //   console.log('test method');
  //     // }
  //     console.log((IE.document).fafa);
  //     await sleep(1000);
  //   }
  // } catch(err) {
  //   console.log(err);
  // }
  // // IE.Quit();
  // console.log('end');
// })();
