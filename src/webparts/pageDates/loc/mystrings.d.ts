declare interface IPageDatesWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ShowDatesFieldLabel: string;
  ShowDatesFieldLabelOptionAuto: string;
  ShowDatesFieldLabelOptionCreated: string;
  ShowDatesFieldLabelOptionModified: string;
  ShowDatesFieldLabelOptionBoth: string;
  AdjustmentsFieldGroupName: string;
  PublishPageFirst: string;
  EditPublishDateTimeFieldLabel: string;
  ManualModifiedDateFieldLabel: string;
  UseAutoModifiedDateFieldLabel: string;
  Published: string;
  Modified: string;
  ReloadPageToShowPublishDate: string;
  DRAFT: string;
  DateTimeSeparator: string;
  Unpublished: string;
  ExtraTools: string;
  SaveAndUnpublish: string;
}

declare module 'PageDatesWebPartStrings' {
  const strings: IPageDatesWebPartStrings;
  export = strings;
}
