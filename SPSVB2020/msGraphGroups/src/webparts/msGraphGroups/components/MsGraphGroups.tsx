import * as React from 'react';
import styles from './MsGraphGroups.module.scss';
import { IMsGraphGroupsProps } from './IMsGraphGroupsProps';
import { IMsGraphGroupsState } from './IMsGraphGroupsState';
import { escape } from '@microsoft/sp-lodash-subset';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { List } from 'office-ui-fabric-react/lib/List';

export default class MsGraphGroups extends React.Component<IMsGraphGroupsProps, IMsGraphGroupsState> {
  private readonly useSampleData: boolean = true;
  constructor(props: IMsGraphGroupsProps) {
    super(props);

    this.state = {
      groups: []
    };
  }

  public componentDidMount(): void {
    this.props.graphClient
      .api('/me/memberOf')
      .get((error: any, graphResponse: any, rawResponse?: any) => {
        const myGroups: MicrosoftGraph.Group[] = graphResponse.value;
        console.log('MyGroups:', myGroups);
        this.setState({ groups: myGroups });
      });
  }

  private _onRenderEventCell(item: MicrosoftGraph.Group, index: number | undefined): JSX.Element {
    return (
      <div>
        <h3>{item.displayName}</h3>
        <h4>{item.description}</h4>
      </div>
    );
  }

  public render(): React.ReactElement<IMsGraphGroupsProps> {
    return (
      <List items={this.state.groups} onRenderCell={this._onRenderEventCell} />
    );
  }
}
