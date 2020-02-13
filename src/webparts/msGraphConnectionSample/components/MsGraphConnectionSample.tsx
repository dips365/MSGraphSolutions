import * as React from 'react';
import styles from './MsGraphConnectionSample.module.scss';
import { IMsGraphConnectionSampleProps } from './IMsGraphConnectionSampleProps';
import { IMsGraphConnectionSampleState } from "./IMsGraphConnectionSampleState";
import { escape } from '@microsoft/sp-lodash-subset';
import {MSGraphClient} from '@microsoft/sp-http';
import { IUserInformation } from "./IUserInformation";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import {
  autobind,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode
} from 'office-ui-fabric-react';
import { ServiceKey } from '@microsoft/sp-core-library';

// Configure the columns for the DetailsList component
let _usersListColumns = [
  {
    key: 'displayName',
    name: 'Display name',
    fieldName: 'displayName',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  }
  // {
  //   key: 'email',
  //   name: 'email',
  //   fieldName: 'email',
  //   minWidth: 50,
  //   maxWidth: 100,
  //   isResizable: true
  // },
  // {
  //   key: 'userPrincipalName',
  //   name: 'User Principal Name',
  //   fieldName: 'userPrincipalName',
  //   minWidth: 100,
  //   maxWidth: 200,
  //   isResizable: true
  // },
];

export default class MsGraphConnectionSample extends React.Component<IMsGraphConnectionSampleProps, IMsGraphConnectionSampleState> {
  constructor(props:IMsGraphConnectionSampleProps,sttate:IMsGraphConnectionSampleState){
    super(props);

    this.state = {
      users:[]
    };
  }
  public render(): React.ReactElement<IMsGraphConnectionSampleProps> {

    return (
      <div className={ styles.msGraphConnectionSample }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
              <p>
              <PrimaryButton
                    text='Get Details'
                    title='Get Details'
                    onClick={ this.GetUserProfileInformation }
                  />
              </p>
              {
                (this.state.users != null && this.state.users.length > 0) ?
                  <p className={ styles.form }>
                  <DetailsList
                    items={ this.state.users }
                    columns={ _usersListColumns }
                    setKey='set'
                    checkboxVisibility={ CheckboxVisibility.hidden }
                    selectionMode={ SelectionMode.none }
                    layoutMode={ DetailsListLayoutMode.fixedColumns }
                    compact={ false }
                  />
                </p>
                : null
              }
            </div>
          </div>
        </div>
      </div>
    );
  }
  private SampleClick():void {
    alert("sample");
  }

  @autobind
  private GetUserProfileInformation():void {
    try {

      this.props.context.msGraphClientFactory
      .getClient().then((client:MSGraphClient):void=>{
          client.api("Users")
          .version('v1.0')
          .select("displayName,mail,userPrincipalName")
          .get((error, response: any, rawResponse?: any) => {
            if (error) {
              console.error(error);
              return;
            }

          // Prepare the output array
          var users: Array<IUserInformation> = new Array<IUserInformation>();

           // Map the JSON response to the output array
          response.value.map((item: any) => {
          users.push({
            displayName: item.displayName,
            email: item.mail,
            userPrincipalName: item.userPrincipalName,
          });
          // Update the component state accordingly to the result
          this.setState(
          {
            users: users,
          });

        });

        });
      });
    } catch (error) {
      throw error;
    }
  }
}
