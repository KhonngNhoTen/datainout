export interface ISheetMeta {
  get IsCompleted(): boolean;
  get RowCount(): number;
  updateRowCount(isCompleted: boolean, rowCount?: number): void;
  completeJob(index: number): void;
  getSheetName(index?: number): string;
  getSheetStatus(name: string): boolean;
}
export type SheetMetaOptions = {
  name: string;
  maxRow?: number;
  jobs?: number[];
};
export class SheetMeta {
  private byJobIndex: boolean = true;
  private sheetMetas: Record<
    string,
    {
      maxRow?: number;
      jobCount?: number;
      rowCount: number;
      isCompleted: boolean;
    }
  > = {};
  private mapSheetByJobIndex: Record<number, string> = {};
  private currentSheetName: number = 0;
  private rowCount: number = 0;
  private sheetNames: string[] = [];
  private isCompleted: boolean = false;
  private jobCount: number = 0;

  constructor(opts: SheetMetaOptions[]) {
    for (let i = 0; i < opts.length; i++)
      if (opts[i].jobs || opts[i].maxRow) this.addSheetMeta(opts[i].name, opts[i].jobs ?? opts[i].maxRow);
  }

  // private addSheetMeta(name: string, jobIndexes: number[]): this;
  // private addSheetMeta(name: string, rowCount: number): this;
  private addSheetMeta(name: string, data: unknown): this {
    if (typeof data === "number") {
      this.byJobIndex = false;
      this.sheetMetas[name] = {
        isCompleted: false,
        rowCount: 0,
        maxRow: data,
      };
    } else if (Array.isArray(data)) {
      this.byJobIndex = true;
      this.sheetMetas[name] = {
        isCompleted: false,
        rowCount: 0,
        jobCount: data.length,
        maxRow: 0,
      };
      this.jobCount += data.length;
      data.forEach((e) => {
        this.mapSheetByJobIndex[e] = name;
      });
    }
    this.sheetNames.push(name);
    return this;
  }

  private completeJob(index: number) {
    const name = this.mapSheetByJobIndex[index];
    if (this.sheetMetas[name].jobCount) {
      this.sheetMetas[name].jobCount -= 1;
      this.sheetMetas[name].isCompleted = this.sheetMetas[name].jobCount <= 0;
      this.jobCount--;
      this.isCompleted = this.jobCount <= 0;
    }
  }

  public get IsCompleted(): boolean {
    return this.isCompleted;
  }

  public get RowCount(): number {
    return this.rowCount;
  }

  private updateRowCount(isCompleted: boolean, rowCount: number) {
    const name = this.sheetNames[this.currentSheetName];
    const sheetMeta = this.sheetMetas[name];
    // if (!rowCount) {
    //   this.sheetMetas[name].isCompleted = true;
    //   this.isCompleted = true;
    // } else {
    if (isCompleted) sheetMeta.isCompleted = isCompleted;
    else if (!this.byJobIndex) {
      this.rowCount += rowCount;
      if ((sheetMeta.maxRow ?? 0) < sheetMeta.rowCount) {
        this.currentSheetName++;
        sheetMeta.isCompleted = true;
      }
      sheetMeta.rowCount += rowCount;
    }
    // }
  }

  private getSheetName(index?: number) {
    if (this.byJobIndex && index) return this.mapSheetByJobIndex[index];
    return this.sheetNames[this.currentSheetName];
  }

  private getSheetStatus(name: string) {
    return this.sheetMetas[name].isCompleted;
  }
}
