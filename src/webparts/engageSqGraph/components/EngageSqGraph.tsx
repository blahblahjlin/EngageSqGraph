import * as React from 'react';
import styles from './EngageSqGraph.module.scss';
import { IEngageSqGraphProps } from './IEngageSqGraphProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as moment from 'moment';

// ***********

import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardTitle,
  DocumentCardLogo,
  DocumentCardStatus,
  IDocumentCardLogoProps,
  IDocumentCardActivityPerson,
  IDocumentCardStyles,
} from '@fluentui/react/lib/DocumentCard';
import { mergeStyles } from '@fluentui/react/lib/Styling';
import {ActivityItem, Icon, mergeStyleSets, DEFAULT_ROW_HEIGHTS, MessageBar, MessageBarType,MessageBarButton, Link, themeRulesStandardCreator, personaPresenceSize, FabricSlots } from 'office-ui-fabric-react';

const conversationTileClass = mergeStyles({ height: 182 });

const logoProps: IDocumentCardLogoProps = {
  logoIcon: 'OutlookLogo',
};

const cardStyles: IDocumentCardStyles = {
  root: { display: 'inline-block', marginRight: 20, width: 320 },
};


const EmailInfoBar = () => (
  <MessageBar      
    messageBarType={MessageBarType.warning}
    isMultiline={false}
    
  >
    Latest 5 emails..
  </MessageBar>
);

// ***********

export interface EmailContents {
  subject: string;
  authorEmail: string;
  body: string;
  createdDate: string;
  authorPerson? :IDocumentCardActivityPerson[];
  emailWebLink : string;
}

export interface IStates {
  UnreadMail: EmailContents[]

}

export default class EngageSqGraph extends React.Component<IEngageSqGraphProps, IStates> {

  public _unreadMail: EmailContents[] = [];

  constructor(props:any) {
    super(props);

    //this.setState({UnreadMail: []});
    this.state = {UnreadMail: []};
    this.GetEmailData();
  }

  public GetEmailData(){

    let unreadEmails: EmailContents[] = [];

    this.props.spcontext.msGraphClientFactory
    .getClient('3')
    .then((client: MSGraphClientV3): void => {
      // get information about the current user from the Microsoft Graph
      client
      .api('/me/mailFolders/Inbox/messages')
      .top(5)
      .orderby("receivedDateTime desc")
      .get((error, messages: any, rawResponse?: any) => {
        
        for (let index = 0; index < messages.value.length; index++) {
          let emailItem = {} as EmailContents;
          
     
          emailItem.authorEmail = messages.value[index].from.emailAddress.address;
          emailItem.subject = messages.value[index].subject;
          emailItem.body = messages.value[index].bodyPreview;
          emailItem.createdDate = messages.value[index].createdDateTime;

          let senderName = messages.value[index].sender.emailAddress.name;
          //let initials = senderName.shift().charAt(0) + senderName.pop().charAt(0);
          var names = senderName.split(' '), initials = names[0].substring(0, 1).toUpperCase();

          if (names.length > 1) {
            initials += names[names.length - 1].substring(0, 1).toUpperCase();
        }
          emailItem.authorPerson = [];
          emailItem.authorPerson.push({ name: senderName, profileImageSrc: '', initials: initials });
          emailItem.emailWebLink = messages.value[index].webLink;
        
          unreadEmails.push(emailItem);
        }

        this.setState({UnreadMail: unreadEmails});

      });
    });
  }

  public render(): React.ReactElement<IEngageSqGraphProps> {

    //this.GetEmailData();

    const {
      userDisplayName,
      currentUserEmail,
      currentUserJobTitle,
      currentUserOfficeLocation,

      spcontext
    } = this.props;

    return (
      <section className="headerES">
        <div>
          Display name: {this.props.userDisplayName}<br/>
          Email: {this.props.currentUserEmail}<br/>
          JobTitle: {this.props.currentUserJobTitle} <br/>
          Office Location {this.props.currentUserOfficeLocation}<br/>
        <br/>
        <EmailInfoBar/>
        <br/>
          {
           this.state.UnreadMail.map(email => {
            return (
            //  <p>Subject: {email.subject}</p>
            <DocumentCard
                aria-label={
                  'blahblahblah.'
                }
                styles={cardStyles}
                onClickHref={email.emailWebLink}
                onClickTarget="_blank"
              >
                <DocumentCardLogo {...logoProps} />
                <div className={conversationTileClass}>
                  <DocumentCardTitle title={email.subject} shouldTruncate />
                  <DocumentCardTitle
                    title={email.body}
                    shouldTruncate
                    showAsSecondaryTitle
                  />
                </div>
                <DocumentCardActivity activity={moment.utc(email.createdDate).local().format("DD/MM/YYYY hh:mm a")} people={email.authorPerson.slice(0,1)} />
                
            </DocumentCard>

            );
          })
          }

          <br/>


        </div>
      </section>
    );
  }
}
