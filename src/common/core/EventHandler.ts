import { EventType } from "../types/common-type.js";

class ContextEventHook {
  private context: any = {};
  after(callback: (ctx: any) => void) {
    callback(this.context);
  }
  before(callback: (ctx: any) => void): this {
    callback(this.context);
    return this;
  }
}

export class ContextEventHandler {
  private groupContext: any = {};
  private setEvent(key: string) {
    const eventHook = new ContextEventHook();
    if (!this.groupContext[key]) this.groupContext[key] = [];
    this.groupContext[key].push(eventHook);
    return eventHook;
  }
  get WorkBook(): ContextEventHook {
    return this.setEvent("workBook");
  }
  get Sheet(): ContextEventHook {
    return this.setEvent("sheet");
  }
  get Row(): ContextEventHook {
    return this.setEvent("row");
  }
  get Header(): ContextEventHook {
    return this.setEvent("header");
  }
  get Footer(): ContextEventHook {
    return this.setEvent("footer");
  }
  get Error(): ContextEventHook {
    return this.setEvent("error");
  }
}

export abstract class HookProvider {
  abstract do(context: ContextEventHandler): void;
}

export type EventHandlerType = {
  workBook: () => void;
  endWorkBook: () => void;

  sheet: (sheetName?: string) => void;
  endSheet: (sheetName?: string) => void;

  onRow: () => void;
  endRow: () => void;

  header: (sheetName?: string) => void;
  endHeader: (sheetName?: string) => void;

  footer: (sheetName?: string) => void;
  endFooter: (sheetName?: string) => void;
  /** Handle error. Return false to cancel import, otherhands return true */
  error: (error: Error) => boolean;
};

export class EventHandler {
  private listEvents: Partial<EventHandlerType> = {};

  onFile(func: EventHandlerType["workBook"]) {
    this.listEvents.workBook = func;
  }
  endFile(func: EventHandlerType["endWorkBook"]) {
    this.listEvents.endWorkBook = func;
  }

  begin(func: EventHandlerType["sheet"]) {
    this.listEvents.sheet = func;
  }
  end(func: EventHandlerType["endSheet"]) {
    this.listEvents.endSheet = func;
  }

  error(func: EventHandlerType["error"]) {
    this.listEvents.error = func;
  }

  private buildHook(): HookProvider {
    const that = this;
    return new (class extends HookProvider {
      do(context: ContextEventHandler): void {
        const workBookFunc = that.listEvents.workBook;
        const endWorkBookFunc = that.listEvents.endWorkBook;
        const sheetFunc = that.listEvents.sheet;
        const endSheetFunc = that.listEvents.endSheet;
        const errorFunc = that.listEvents.error;

        if (workBookFunc) context.WorkBook.before((ctx) => workBookFunc());
        if (endWorkBookFunc) context.WorkBook.after((ctx) => endWorkBookFunc());

        if (sheetFunc) context.Sheet.before((ctx) => sheetFunc(ctx.sheetName));
        if (endSheetFunc) context.Sheet.after((ctx) => endSheetFunc(ctx.sheetName));

        if (errorFunc) context.Error.after((ctx) => errorFunc(ctx.error));
      }
    })();
  }
}
