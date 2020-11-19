import { CrebassObject } from "./interface";

declare global {
  interface Document {
    __crebass__: CrebassObject;
    __crebass__init: boolean;
  }
}

if (document.__crebass__ && !document.__crebass__init) {
  console.log('prepare to crebass install');
  document.__crebass__.log('INSTALL', 'prepare to crebass install');
  document.__crebass__init = true;
}
