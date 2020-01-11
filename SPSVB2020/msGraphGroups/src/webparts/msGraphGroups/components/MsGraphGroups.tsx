import * as React from 'react';
import styles from './MsGraphGroups.module.scss';
import { IMsGraphGroupsProps } from './IMsGraphGroupsProps';
import { IMsGraphGroupsState } from './IMsGraphGroupsState';
import { escape } from '@microsoft/sp-lodash-subset';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { List } from 'office-ui-fabric-react/lib/List';

export default class MsGraphGroups extends React.Component<IMsGraphGroupsProps, IMsGraphGroupsState> {
  constructor(props: IMsGraphGroupsProps) {
    super(props);

    this.state = {
      groups: []
    };
  }

  public componentDidMount(): void {
    this.props.graphClient
      .api('/me/memberOf')
      .select("displayName,description,mail,groupTypes,visibility")
      .get((error: any, graphResponse: any, rawResponse?: any) => {
        const myGroups: MicrosoftGraph.Group[] = graphResponse.value;
        console.log('MyGroups:', myGroups);
        this.setState({ groups: myGroups });
      });
  }

  private _onRenderEventCell(item: MicrosoftGraph.Group, index: number | undefined): JSX.Element {
    return (
      <div>
        <h3>{item.displayName} ({item.mail})</h3>
        <p>{item.description}</p>
      </div>
    );
  }

  public render(): React.ReactElement<IMsGraphGroupsProps> {
    return (
      <div>
        <h1>My Groups</h1>
        <List items={this.state.groups} onRenderCell={this._onRenderEventCell} />
      </div>
    );
  }
}
