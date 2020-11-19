export interface CrebassObject {
  log: (track: string, content: string) => void;
  logJson: (track: string, json: string) => void;
  onerror: (message: string, URL: string, line: number, column: number, error: Error) => void;
  testMethod: (a: any, b: any, c: any, d: any, e: any) => void;
  ping: () => number;
  resolve: (id: number, result: any) => void;
}
