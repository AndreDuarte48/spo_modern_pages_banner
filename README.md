# Banner web part

This web part provides you the ability to add a variable height image banner with a linkable title.
**note** As of V2, this Banner webpart is hidden from the webpart toolbox and may be only added to a webpart using PowerShell or another automated deployment approach. This may be overwritten by updating the hiddenFromToolbox property to "false" found in [./src/webparts/banner/BannerWebPart.manifest.json](./src/webparts/banner/BannerWebPart.manifest.json).


## Configurable Properties

The `Banner` webpart can be configured with the following properties:

| Label | Property | Type | Required | Description |
| ---- | ---- | ---- | ---- | ---- |
| Overlay image text | bannerText | string | no | The text message or title you want displayed on the banner image |
| Secondary text | bannerSecondaryText | string | no | The text message or title you want displayed below the banner title |
| Image URL | bannerImage | string | no | The url of the banner image |
| Link URL | bannerLink | string | no | The hyperlink url of the bannerText link |
| Banner height | bannerHeight | number | no | Provides the fixed height of the banner image |
| Enable parallax effect | useParallax | toggle | no | Enable if you want to include parallax effect on vertical scrolling |

## How to use this web part on your web pages

1. Place the page you want to add this web part to in edit mode.
2. Search for and insert the **Banner** web part.
3. Configure the web part to update its properties.

## Used SharePoint Framework Version

![drop](https://img.shields.io/badge/version-1.16.1-green.svg)

* Supported in SharePoint Online

## Applies to

* [SharePoint Framework](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
* [Office 365 tenant](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)

## Prerequisites

none

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

* Clone this repository
* in the command line run:
  * `npm install`
  * `gulp serve`

> Include any additional steps as needed.
