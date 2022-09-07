import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneButton,
} from '@microsoft/sp-property-pane';

import { PropertyFieldDateTimePicker, DateConvention } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
import { IDateTimeFieldValue } from "@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker";

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PageDatesWebPartStrings';
import PageDates from './components/PageDates';
import { IPageDatesProps, ShowDates } from './components/IPageDatesProps';

import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/items/list";
import "@pnp/sp/files/web";
import { IFile } from '@pnp/sp/files/types';

interface IPageDatesWebPartProps {
  manualCreatedDate: IDateTimeFieldValue;
  manualModifiedDate: IDateTimeFieldValue;
  showDates: ShowDates;
}

interface ISitePageDates {
  created: Date;
  modified: Date;
  firstPublished?: Date;
}

interface IPageDates {
  Created: string;
  Modified: string;
  FirstPublishedDate?: string;
}

interface IPageVersions {
  MinorVersion: number;
  MajorVersion: number;
}

export default class PageDatesWebPart extends BaseClientSideWebPart<IPageDatesWebPartProps> {

  private _file?: IFile;
  private _dates?: ISitePageDates;
  private _isNew = true;
  private _isDraft = false;
  private _isDarkTheme: boolean = false;
  private _unpublishButtonPressed = false;
  private _sp: SPFI;

  protected async onInit(): Promise<void> {
    await super.onInit();

    this._sp = spfi().using(SPFx(this.context));
    await this._updateContext();
  }

  public async render(): Promise<void> {
    await this._updateContext();
    const {manualCreatedDate, manualModifiedDate, showDates} = this.properties;
    const element: React.ReactElement<IPageDatesProps> = React.createElement(
      PageDates,   
      {
        publishedDate: manualCreatedDate && manualCreatedDate.value
          ? new Date(manualCreatedDate.value as unknown as React.ReactText)
          : this._dates && !this._isNew ? this._dates.firstPublished || this._dates.created : undefined,
        modifiedDate: manualModifiedDate && manualModifiedDate.value
          ? new Date(manualModifiedDate.value as unknown as React.ReactText)
          : this._dates ? this._dates.modified : undefined,
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        displayMode: this.displayMode,
        showDates,
        isDraft: this._isDraft,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    await this._updateContext(true);
    super.onPropertyPaneConfigurationStart();
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
              groupFields: [
                PropertyPaneChoiceGroup('showDates', {
                  label: strings.ShowDatesFieldLabel,
                  options: [
                    { key: ShowDates.Auto, text: strings.ShowDatesFieldLabelOptionAuto },
                    { key: ShowDates.Created, text: strings.ShowDatesFieldLabelOptionCreated },
                    { key: ShowDates.Modified, text: strings.ShowDatesFieldLabelOptionModified },
                    { key: ShowDates.Both, text: strings.ShowDatesFieldLabelOptionBoth },
                  ],
                }),
              ],
            },
            {
              groupName: `${strings.AdjustmentsFieldGroupName}${this._isNew ? ` (${strings.PublishPageFirst})` : ''}`,
              groupFields: [
                PropertyFieldDateTimePicker('manualCreatedDate', {
                  label:  strings.EditPublishDateTimeFieldLabel,
                  disabled: this._isNew,
                  initialDate: this.properties.manualCreatedDate
                    || (this._dates && this._dateToDateField(this._dates.firstPublished))
                    || (this._dates && this._dateToDateField(this._dates.created)),
                  dateConvention: DateConvention.DateTime,
                  onPropertyChange: async (propertyPath, oldValue, newValue) => {
                    const newDate: Date = newValue.value;
                    const item = await this._file.getItem();
                    await item.validateUpdateListItem([{
                      FieldName: "FirstPublishedDate",
                      FieldValue: `${newDate.toLocaleDateString()} ${newDate.toLocaleTimeString()}`
                    }]);
                    this.onPropertyPaneFieldChanged(propertyPath, oldValue, false);
                  },
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'manualCreatedDate',
                  showLabels: false
                }),
                PropertyFieldDateTimePicker('manualModifiedDate', {
                  label: strings.ManualModifiedDateFieldLabel,
                  disabled: this._isNew,
                  initialDate: this.properties.manualModifiedDate,
                  dateConvention: DateConvention.DateTime,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'manualModifiedDate',
                  showLabels: false
                }),
                PropertyPaneButton('manualModifiedDate', {
                  text: `${strings.UseAutoModifiedDateFieldLabel}${this._dates && this._dates.modified && ` (${this._dates.modified.toLocaleDateString()})`}`,
                  disabled: this._isNew,
                  onClick: (value: unknown) => {
                    this.properties.manualModifiedDate = null;
                    this.context.propertyPane.close();
                    this.context.propertyPane.open();
                  },
                }),
              ],
            },
            {
              groupName: strings.ExtraTools,
              groupFields: [
                PropertyPaneButton('unpublish',{
                  text: strings.SaveAndUnpublish,
                  disabled: this._isNew || this._unpublishButtonPressed,
                  onClick: async () => {
                    await this._file.checkin();
                    await this._file.unpublish(strings.Unpublished);
                    await this._file.checkout();
                    this._unpublishButtonPressed = true;
                    this.context.propertyPane.close();
                    this.context.propertyPane.open();
                  },
                }),
              ],
            },
        
          ]
        }
      ]
    };
  }

  private async _updateContext(force?: boolean): Promise<void> {
    if (this._dates && force !== true) return;
    try {
      const {list: {id: listId}, listItem: {id: listItemId } } = this.context.pageContext;
      const dateFields: IPageDates = await this._sp.web.lists.getById(listId.toString()).items.getById(listItemId).select("Modified", "Created","FirstPublishedDate")();
      // @ts-expect-error: In this context, listItem actually has a uniqueId property
      this._file = this._sp.web.getFileById(this.context.pageContext.listItem.uniqueId);
      const pageVersions: IPageVersions = await this._file.select("MinorVersion", "MajorVersion")();
      this._isNew = pageVersions.MajorVersion === 0;
      this._isDraft = pageVersions.MinorVersion !== 0;
      this._dates = {
        created: new Date(dateFields.Created),
        modified: new Date(dateFields.Modified),
        firstPublished: dateFields.FirstPublishedDate && new Date(dateFields.FirstPublishedDate),
      };
    } catch(e) {
      console.error(e)
    }
  }

  private _dateToDateField(date: Date): IDateTimeFieldValue | undefined {
    if (date) return {
      value: date,
      displayValue: date.toLocaleString(),
    };
  }

}
