import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EngageSqGraphWebPartStrings';
import EngageSqGraph from './components/EngageSqGraph';
import { IEngageSqGraphProps } from './components/IEngageSqGraphProps';

import { GraphFI, graphfi, SPFx } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/messages";
import "@pnp/graph/onedrive";

import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { TeachingBubbleBase } from 'office-ui-fabric-react';
import { Message } from '@pnp/graph/messages';
import { _Drive, _DriveItem, _DriveItems } from '@pnp/graph/onedrive/types';

export interface IEngageSqGraphWebPartProps {
  description: string;
  currentUserEmail:string;
}

export interface EmailContents {
  subject: string;
  authorEmail: string;
  body: string;
  createdDate: string;
}

export default class EngageSqGraphWebPart extends BaseClientSideWebPart<IEngageSqGraphWebPartProps> {

  private _currentUserDisplayName: string = "";
  private _currentUserEmail: string = "";
  private _currentUserJobTitle: string ="";
  private _currentUserOfficeLocation: string = "";
  private _unreadMail: EmailContents[] = [];


  private _driveItems: _Drive;

  public render(): void {
    const element: React.ReactElement<IEngageSqGraphProps> = React.createElement(
      EngageSqGraph,
      {
        userDisplayName: this._currentUserDisplayName,

        currentUserEmail: this._currentUserEmail,
        currentUserJobTitle: this._currentUserJobTitle,
        currentUserOfficeLocation: this._currentUserOfficeLocation,
        unreadEmail: this._unreadMail,
        spcontext: this.context

      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    //return super.onInit();
    await super.onInit();
    const graph = graphfi().using(SPFx(this.context));


    // call graph to get user data      
    let currentUserData = await graph.me();
    this._currentUserEmail = currentUserData.mail;
    this._currentUserDisplayName = currentUserData.displayName;
    this._currentUserJobTitle = currentUserData.jobTitle;
    this._currentUserOfficeLocation = currentUserData.officeLocation;

  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
