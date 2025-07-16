type ParamsPartialDataCallback = {
  items: any[] | null;
  hasNext: boolean;
  sheetName: string;
  jobIndex: number;
  sheetStatus: "completed" | "running";
  status: "completed" | "running";
};
type PartialDataCallback = (params: ParamsPartialDataCallback) => Promise<any>;
export abstract class PartialDataTransfer {
  private callback: PartialDataCallback = async () => {};
  private delayMs: number = 10;
  private sheetStatuses: { [sheetname: string]: { status: boolean; jobStatuses: boolean[] } } = {};

  protected jobCount = 1;

  constructor(jobCount?: number, delayMs?: number) {
    this.delayMs = delayMs ?? 10;
    if (jobCount && jobCount < 0) throw new Error("jobCount must be greater than 0");
    this.jobCount = jobCount ?? 1;
  }

  private async runJob(jobIndex: number, originalSheetName: string, callback: PartialDataCallback) {
    this.callback = callback;
    let isLoop = true;
    while (isLoop) {
      let sheetName = originalSheetName;
      const { items, hasNext } = await this.fetchBatch(jobIndex);
      if (this.jobCount > 1) sheetName = this.bindJob2Sheet(jobIndex, originalSheetName) ?? originalSheetName;

      if (!this.sheetStatuses[sheetName]) this.sheetStatuses[sheetName] = { jobStatuses: [], status: false };
      this.sheetStatuses[sheetName].jobStatuses[jobIndex] = !hasNext;
      this.sheetStatuses[sheetName].status = this.isSheetStatus(sheetName);

      await this.callback({
        items,
        hasNext,
        sheetName,
        jobIndex,
        sheetStatus: this.sheetStatuses[sheetName].status ? "completed" : "running",
        status: this.isStatus(),
      });

      await new Promise((resolve) => setTimeout(resolve, this.delayMs));
      isLoop = hasNext;
    }
  }

  private isStatus() {
    const sheetNames = Object.keys(this.sheetStatuses);
    for (let i = 0; i < sheetNames.length; i++) {
      if (this.sheetStatuses[sheetNames[i]].status === false) return "running";
    }
    return "completed";
  }

  /** True is completed , False is running */
  private isSheetStatus(sheetName: string) {
    return !this.sheetStatuses[sheetName].jobStatuses.includes(false);
  }

  async run(sheetName: string, callback: PartialDataCallback) {
    const promises: any[] = [];
    const that = this;
    if (this.jobCount > 1) {
      for (let i = 0; i < this.jobCount; i++) promises.push(that.runJob(i, sheetName, callback));
      await Promise.all(promises);
    } else await this.runJob(0, sheetName, callback);
  }

  protected bindJob2Sheet(jobIndex: number, originalSheetName: string): null | string {
    return null;
  }

  abstract fetchBatch(jobIndex: number): Promise<{ items: any[] | null; hasNext: boolean }>;
}
