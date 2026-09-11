import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'AdvanceSearchWebPartStrings';
import AdvanceSearch from './components/AdvanceSearch';
import { IAdvanceSearchProps } from './components/IAdvanceSearchProps';
import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import "@pnp/sp/search";

export interface IAdvanceSearchWebPartProps {
  description: string;
  urldata: string;
}
export let sp: SPFI;
export default class AdvanceSearchWebPart extends BaseClientSideWebPart<IAdvanceSearchWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _resizeObserver: ResizeObserver | null = null;
  private _resizeHandler: (() => void) | null = null;
  private _rafId: number | null = null;

  private _scheduleFullBleed(): void {
    if (this._rafId !== null) {
      cancelAnimationFrame(this._rafId);
    }
    this._rafId = requestAnimationFrame(() => {
      this._applyFullBleed();
      this._rafId = null;
    });
  }

  private _applyFullBleed(): void {
    if (!this.domElement) {
      return;
    }

    this._ensureFullBleedStyleTag();

    this.domElement.classList.add('spfx-full-bleed-webpart');
    this.domElement.style.setProperty('width', '100%', 'important');
    this.domElement.style.setProperty('max-width', 'none', 'important');
    this.domElement.style.setProperty('padding', '0', 'important');
    this.domElement.style.setProperty('margin', '0', 'important');

    let el: HTMLElement | null = this.domElement.parentElement;
    while (el && el !== document.body) {
      el.classList.add('spfx-full-bleed-ancestor');
      el.style.setProperty('max-width', 'none', 'important');
      el.style.setProperty('width', '100%', 'important');
      el.style.setProperty('padding-left', '0', 'important');
      el.style.setProperty('padding-right', '0', 'important');
      el.style.setProperty('margin-left', '0', 'important');
      el.style.setProperty('margin-right', '0', 'important');

      if (el.id === 'spPageCanvasContent' || el.getAttribute('data-automation-id') === 'Canvas') {
        break;
      }
      el = el.parentElement;
    }
  }

  private _ensureFullBleedStyleTag(): void {
    const styleId = 'spfx-advancesearch-team-site-fullbleed';
    if (!document.getElementById(styleId)) {
      const styleEl = document.createElement('style');
      styleEl.id = styleId;
      styleEl.textContent = `
        .spfx-full-bleed-ancestor,
        .CanvasZoneContainer:has(.spfx-full-bleed-webpart),
        .CanvasZone:has(.spfx-full-bleed-webpart),
        .CanvasSection:has(.spfx-full-bleed-webpart),
        .CanvasSection-col:has(.spfx-full-bleed-webpart),
        [class*="CanvasSection-col"]:has(.spfx-full-bleed-webpart),
        [class*="CanvasSection-xl"]:has(.spfx-full-bleed-webpart),
        .ControlZone:has(.spfx-full-bleed-webpart),
        [data-automation-id="CanvasControl"]:has(.spfx-full-bleed-webpart),
        .CanvasZoneContainer,
        [data-automation-id="CanvasZoneContainer"],
        .CanvasZoneContainer--read {
          max-width: none !important;
          width: 100% !important;
          margin-left: 0 !important;
          margin-right: 0 !important;
          padding-left: 0 !important;
          padding-right: 0 !important;
          box-sizing: border-box !important;
        }
        #spPageCanvasContent,
        [data-automation-id="Canvas"],
        [data-automation-id="CanvasLayout"] {
          max-width: none !important;
          width: 100% !important;
        }
      `;
      document.head.appendChild(styleEl);
    }
  }

  public render(): void {
    // Force web part container and SharePoint canvas elements to full bleed on Team Sites and Communication Sites
    this._applyFullBleed();

    if (!this._resizeHandler) {
      this._resizeHandler = () => this._scheduleFullBleed();
      window.addEventListener('resize', this._resizeHandler);
    }

    if (!this._resizeObserver && typeof ResizeObserver !== 'undefined') {
      const target = document.getElementById('spPageCanvasContent') || document.body;
      this._resizeObserver = new ResizeObserver(() => {
        this._scheduleFullBleed();
      });
      this._resizeObserver.observe(target);
    }

    const element: React.ReactElement<IAdvanceSearchProps> = React.createElement(
      AdvanceSearch,
      {
        description: this.properties.description,
        urlSite: this.properties.urldata,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  //protected onInit(): Promise<void> {
  public async onInit(): Promise<void> {

    await super.onInit();

    const initialWebUrl = this.properties.urldata?.trim() || this.context.pageContext.web.absoluteUrl;
    sp = spfi(initialWebUrl).using(SPFx(this.context));
    // Initialize once globally for this web

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
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
    if (this._rafId !== null) {
      cancelAnimationFrame(this._rafId);
      this._rafId = null;
    }
    if (this._resizeObserver) {
      this._resizeObserver.disconnect();
      this._resizeObserver = null;
    }
    if (this._resizeHandler) {
      window.removeEventListener('resize', this._resizeHandler);
      this._resizeHandler = null;
    }
    const styleEl = document.getElementById('spfx-advancesearch-team-site-fullbleed');
    if (styleEl && styleEl.parentNode) {
      styleEl.parentNode.removeChild(styleEl);
    }
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
                }),
                PropertyPaneTextField('urldata', {
                  label: strings.UrlDataFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
