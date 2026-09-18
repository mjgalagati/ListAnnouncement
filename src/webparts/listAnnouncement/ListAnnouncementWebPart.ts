import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ListAnnouncementWebPartStrings';
import ListAnnouncements from './components/ListAnnouncements';
import { IListAnnouncementsProps } from './components/IListAnnouncementsProps';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";

export interface IListAnnouncementWebPartProps {
  webpartTitle: string;
  sourceList: string;
  highlightTypes: string;
  editorGroup: string;
  approverGroup: string;
  hasApproval: boolean;
  itemLimit: number;
  reactionsListName: string;
  commentsListName: string;
  commentsReactionsAlias: string;
}

export default class ListAnnouncementWebPart extends BaseClientSideWebPart<IListAnnouncementWebPartProps> {

  private _sp!: ReturnType<typeof spfi>;
  private _siteLists: { key: string; text: string }[] = [];
  private _currentUserId: number = 0;
  private _isEditor: boolean = false;
  private _isApprover: boolean = false;

  public render(): void {
    const element: React.ReactElement<IListAnnouncementsProps> = React.createElement(
      ListAnnouncements,
      {
        // No default — an empty title means no heading is rendered and no space is reserved for it.
        webpartTitle: this.properties.webpartTitle || '',
        sourceList: this.properties.sourceList,
        highlightTypes: this.properties.highlightTypes || '',
        isDarkTheme: !!this.context.pageContext.legacyPageContext?.isDarkTheme,
        environmentMessage: '',
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        currentUserLogin: this.context.pageContext.user.loginName,
        currentUserId: this._currentUserId,
        isEditor: this._isEditor,
        isApprover: this._isApprover,
        // Undefined on existing/older deployments — default to true so behavior doesn't change until explicitly turned off.
        hasApproval: this.properties.hasApproval !== false,
        itemLimit: this.properties.itemLimit ?? 5,
        reactionsListName: this.properties.reactionsListName || "MThermalReactions",
        commentsListName: this.properties.commentsListName || "MThermalComments",
        commentsReactionsAlias: this.properties.commentsReactionsAlias || "listannouncement",
      } as IListAnnouncementsProps
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context));
    await Promise.all([
      this._loadSiteLists(),
      this._loadCurrentUser(),
      this._checkPermissions(),
    ]);
    return super.onInit();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (newValue !== oldValue) {
      if (propertyPath === 'editorGroup' || propertyPath === 'approverGroup') {
        this._checkPermissions().then(() => this.render()).catch(console.error);
      } else if (propertyPath === 'hasApproval') {
        this._checkPermissions().then(() => {
          this.render();
          this.context.propertyPane.refresh();
        }).catch(console.error);
      } else {
        this.render();
      }
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async _loadCurrentUser(): Promise<void> {
    try {
      const currentUser = await this._sp.web.currentUser.select("Id")();
      this._currentUserId = currentUser.Id;
    } catch (err) {
      console.error("Failed to load current user:", err);
    }
  }

  private async _loadSiteLists(): Promise<void> {
    try {
      const lists = await this._sp.web.lists
        .filter("BaseTemplate eq 100 and Hidden eq false")
        .select("Title")();
      this._siteLists = lists.map(l => ({ key: l.Title, text: l.Title }));
    } catch (err) {
      console.error("Failed to load site lists:", err);
      this._siteLists = [];
    }
  }

  private async _checkPermissions(): Promise<void> {
    try {
      const editorGroupName = this.properties.editorGroup?.trim();
      const approverGroupName = this.properties.approverGroup?.trim();
      if (!editorGroupName && !approverGroupName) {
        this._isEditor = false;
        this._isApprover = false;
        return;
      }
      const userGroups = await this._sp.web.currentUser.groups();
      this._isEditor = !!editorGroupName && userGroups.some(g => g.Title === editorGroupName);
      this._isApprover = !!approverGroupName && userGroups.some(g => g.Title === approverGroupName);
    } catch (err) {
      console.error("Failed to check permissions:", err);
      this._isEditor = false;
      this._isApprover = false;
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('webpartTitle', {
                  label: 'Page Title',
                  description: 'Optional heading shown above the web part. Leave blank to show no heading and reserve no space for it.',
                  placeholder: 'e.g. Announcements',
                }),
                PropertyPaneDropdown('sourceList', {
                  label: 'Select Announcements List',
                  options: this._siteLists,
                  disabled: this._siteLists.length === 0,
                }),
                PropertyPaneTextField('highlightTypes', {
                  label: 'Categories',
                  description: 'Comma-separated (e.g. Safety, Environment, Health)',
                  placeholder: 'Safety, Environment, Health',
                }),
                PropertyPaneTextField('editorGroup', {
                  label: 'Editor Permission Group',
                  description: 'SharePoint site group whose members can create announcements and edit their own. Leave blank to restrict all users.',
                  placeholder: 'e.g. ESH Editors',
                }),
                PropertyPaneToggle('hasApproval', {
                  label: 'Has Approval',
                  onText: 'On',
                  offText: 'Off',
                  checked: this.properties.hasApproval !== false,
                }),
                ...(this.properties.hasApproval !== false ? [
                  PropertyPaneTextField('approverGroup', {
                    label: 'Approver Permission Group',
                    description: 'SharePoint site group whose members can publish, reject, and edit any announcement.',
                    placeholder: 'e.g. ESH Approvers',
                  }),
                ] : []),
                PropertyPaneDropdown('reactionsListName', {
                  label: 'Reactions List',
                  options: this._siteLists,
                  disabled: this._siteLists.length === 0,
                }),
                PropertyPaneDropdown('commentsListName', {
                  label: 'Comments List',
                  options: this._siteLists,
                  disabled: this._siteLists.length === 0,
                }),
                PropertyPaneTextField('commentsReactionsAlias', {
                  label: 'Comments/Reactions Alias',
                  placeholder: 'e.g. listannouncement',
                  description: 'Lowercase key that namespaces this web part\'s rows when the Comments/Reactions lists are shared across multiple web parts.',
                }),
                PropertyPaneSlider('itemLimit', {
                  label: 'Item List Limit',
                  min: 1,
                  max: 5,
                  step: 1,
                  showValue: true,
                  value: this.properties.itemLimit ?? 5,
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
