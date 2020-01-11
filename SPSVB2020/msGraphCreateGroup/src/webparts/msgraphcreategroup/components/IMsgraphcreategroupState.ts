import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IMsgraphcreategroupState {
  response: MicrosoftGraph.ResponseStatus;
  disabled: boolean;
  groupName: string;
  mailName: string;
}
