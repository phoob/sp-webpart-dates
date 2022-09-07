import { DisplayMode } from '@microsoft/sp-core-library';

export enum ShowDates {
  Auto = 'Auto',
  Created = 'Created',
  Modified = 'Modified',
  Both = 'Both',
}

export interface IPageDatesProps {
  publishedDate?: Date;
  modifiedDate?: Date;
  showDates: ShowDates;
  isDraft: boolean;

  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  displayMode: DisplayMode;

}
