export interface IExporter {
  write(...args: any[]): Promise<void>;

  toBuffer(...args: any[]): Promise<Buffer>;

  streamTo(...args: any[]): void;
}
