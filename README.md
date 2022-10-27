# engage-sq-graph

## Summary

Quick demo of the graph solutions

·       How cool it looks and feels (great UI/UX) > to be honest I'm not a UI person so I'm not that great in designing. Give me a mock up however and I'm gold.

·       How useful the component might be - this dashboard is quite useful for us - defense has us locked down pretty tight so to get all the info in one place is great. Given time I will probably tidy it up to include onedrive - recent docs and possibly a search.

·       Integrating Office UI Fabric - done

·       Including a continuous integration pipeline  - I normally do pipelines in azure dev ops > never done it any other way

·       Any interesting and practical software patterns 

·       Use of appropriate PnP libraries - pnp libraries to get basic user info

·       Working as a MS Teams tab and utilising the Teams context - n/a

·       Support for section background theming - n/a

·       Support as a full-bleed web part - done

·       Support dynamic data between web parts - didn't imagine a situation where i'd need this in the current solution. Maybe going forward could split out the basic info, emails and one drive into seperate items and feed it from the basic info wp

·       Integration of other popular and modern React/JavaScript libraries - i've used moment js to tidy up the date display

·       Utilising the Microsoft Graph for operations other than just reading data - was thinking of doing a 'reply email' functionality using https://graph.microsoft.com/v1.0/me/sendMail

·       Utilising the SharePoint REST API to interact with SharePoint data (use of the SP Search API qualifies this) - no I didnt do this > this._context.pageContext.web.absoluteUrl + "/_api/search/query?querytext=" + query > something along those lines 

·       Highly accessible (WCAG AA+) - Didn't know this existed > but no i'd assume mine is D is something with fluent ui holding me up somewhat.

·       Unit tests - nope

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

Uses Graph to Grab user data as well as fetching the top 5 most recent user emails. 

Utilises the fluent ui, moment for date display parsing.

I was going to do the display one drive underneath it but ran out of time. Wife is having a weekend away. Yey for a crying baby for the weekend! wooo

Note: not going to lie I did a rush job on this having dealt with graph in a minimal sense. I did switch from the PNP libraries for graph to extract the emails to MSGraphClientV3. It was for some very strange reason not pushing across any data. so switched to this MSGraphClientV3 and was able to extract it out using: https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/Messages .. got to love graph explorer.
