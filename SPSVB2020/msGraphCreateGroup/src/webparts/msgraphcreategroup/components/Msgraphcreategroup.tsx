import * as React from 'react';
import styles from './Msgraphcreategroup.module.scss';
import { IMsgraphcreategroupProps } from './IMsgraphcreategroupProps';
import { IMsgraphcreategroupState } from './IMsgraphcreategroupState';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrimaryButton, TextField } from 'office-ui-fabric-react';
import { Group } from "@microsoft/microsoft-graph-types";

export default class Msgraphcreategroup extends React.Component<IMsgraphcreategroupProps, IMsgraphcreategroupState> {
  constructor(props: IMsgraphcreategroupProps) {
    super(props);

    this._onClicked = this._onClicked.bind(this);
    this.handleChangeGroupName = this.handleChangeGroupName.bind(this);
    this.handleChangeMailName = this.handleChangeMailName.bind(this);
    this.state = {
      response: null,
      disabled: false,
      groupName: "",
      mailName: ""
    };
  }

  private handleChangeGroupName(event) { this.setState({ groupName: event }); }
  private handleChangeMailName(event) { this.setState({ mailName: event }); }

  private _onClicked(): void {
    const { groupName, mailName } = this.state;
    this.setState({ disabled: true });
    let body: Group = {
      description: "Demo Group for Creating Office 365 Groups via SPFx + MSGraph",
      displayName: groupName,
      groupTypes: [
        "Unified"
      ],
      mailEnabled: true,
      mailNickname: mailName,
      securityEnabled: false
    };
    this.props.graphClient
      .api('/groups')
      .post(body)
      .then((graphResponse) => {
        alert('Created');
        this.setState({ disabled: false, groupName: "", mailName: "" });
      });
  }

  public render(): React.ReactElement<IMsgraphcreategroupProps> {
    const { disabled } = this.state;
    return (
      <div>
        <TextField label="Group Name" onChanged={this.handleChangeGroupName} />
        <TextField label="Mail Name" onChanged={this.handleChangeMailName} />
        <PrimaryButton text="Create Group" onClick={this._onClicked} allowDisabledFocus disabled={disabled} />
      </div>
    );
  }
}
