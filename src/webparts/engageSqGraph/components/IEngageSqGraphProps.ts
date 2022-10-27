import { MSGraphClient } from '@microsoft/sp-http';  
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { _DriveItems } from '@pnp/graph/onedrive/types';


export interface IEngageSqGraphProps {

  userDisplayName: string;
  currentUserEmail: string;
  currentUserJobTitle: string;
  currentUserOfficeLocation: string;


  spcontext: any;
}
