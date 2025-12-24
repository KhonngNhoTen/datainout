export interface ICommandAction {
  handleAction(...data: any): Promise<any>;
}
