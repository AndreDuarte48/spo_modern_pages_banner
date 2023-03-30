import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IWebPartPropertiesMetadata } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from "@microsoft/sp-property-pane";

import { Banner, IBannerProps } from './components';

import * as strings from 'BannerWebPartStrings';
import { getSP } from './pnpjsConfig';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import { SPFI, spfi } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/site-users/web";

export interface IBannerWebPartProps {
  bannerText: string;
  bannerSecondaryText: string;
  bannerImage: string;
  bannerLink: string;
  bannerHeight: number;
  fullWidth: boolean;
  useParallax: boolean;
  useParallaxInt: boolean;
  currentUser: string;
}

export default class BannerWebPart extends BaseClientSideWebPart<IBannerWebPartProps> {
  // tslint:disable-next-line:no-any
  private propertyFieldNumber: any;
  private _environmentMessage = '';
  private _sp: SPFI;

  protected async onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
  
    await super.onInit();

    //Initialize our _sp object that we can then use in other packages without having to pass around the context.
    //  Check out pnpjsConfig.ts for an example of a project setup file.
    getSP(this.context);
    this._sp = getSP();
    await this._getCurrentUser().then((user) => {
      this.properties.currentUser = user.Title;
    });
  }

  public render(): void {
    const element: React.ReactElement<IBannerProps> = React.createElement(
      Banner,
      {
        ...this.properties,
        propertyPane: this.context.propertyPane,
        domElement: this.context.domElement,
        environmentMessage: this._environmentMessage,
        // tslint:disable-next-line:max-line-length
        useParallaxInt: this.displayMode === DisplayMode.Read && !!this.properties.bannerImage && this.properties.useParallax
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Set property metadata
   */
  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'bannerText': { isSearchablePlainText: true },
      'bannerImage': { isImageSource: true },
      'bannerLink': { isLink: true }
    };
  }

  // executes only before property pane is loaded.
  protected async loadPropertyPaneResources(): Promise<void> {
    // import additional controls/components

    const { PropertyFieldNumber } = await import(
      /* webpackChunkName: 'pnp-propcontrols-number' */
      '@pnp/spfx-property-controls/lib/propertyFields/number'
    );

    this.propertyFieldNumber = PropertyFieldNumber;
  }

  /**
   * Property pane configuration
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.BannerConfigName,
              groupFields: [
                PropertyPaneTextField('bannerText', {
                  label:  strings.BannerTextField,
                  multiline: true,
                  maxLength: 200,
                  value: this.properties.bannerText
                }),
                PropertyPaneTextField('bannerImage', {
                  label:  strings.BannerImageUrlField,
                  onGetErrorMessage: this._validateImageField,
                  value: this.properties.bannerImage
                }),
                PropertyPaneTextField('bannerLink', {
                  label:  strings.BannerLinkField,
                  value: this.properties.bannerLink
                }),
                this.propertyFieldNumber('bannerHeight', {
                  key: 'bannerHeight',
                  label:  strings.BannerNumberField,
                  value: this.properties.bannerHeight,
                  maxValue: 500,
                  minValue: 100
                }),
                PropertyPaneToggle('useParallax', {
                  label:  strings.BannerParallaxField,
                  checked: this.properties.useParallax
                }),
                PropertyPaneTextField('bannerSecondaryText', {
                  label:  strings.BannerSecondaryText,
                  value: this.properties.bannerSecondaryText
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  /**
   * Field validation
  */
  private _validateImageField(imgVal: string): string {
    if (imgVal) {
      const urlSplit: string[] = imgVal.split('.');
      if (urlSplit && urlSplit.length > 0) {
        const extName: string = urlSplit.pop().toLowerCase();
        if (['jpg', 'jpeg', 'png', 'gif'].indexOf(extName) === -1) {
          return strings.BannerValidationNotImage;
        }
      }
    }
    return '';
  }
  
  private _getEnvironmentMessage(): string {
    if (!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  private async _getCurrentUser(): Promise<ISiteUserInfo> {
    const sp = spfi(this._sp);
    return await sp.web.currentUser();
  }
}
