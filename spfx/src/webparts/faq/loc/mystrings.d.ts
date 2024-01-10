declare interface IFaqWebPartStrings {
  EnableLoadingFieldLabel: string;
  ListNameFieldDescription: string;
  ListNameFieldLabel: string;
  PaginationLimitFieldLabel: string;
  TitleFieldDescription: string;
  TitleFieldLabel: string;
  ViewNameFieldDescription: string;
  ViewNameFieldLabel: string;
  WebUrlFieldDescription: string;
  WebUrlFieldLabel: string;
}

declare module 'FaqWebPartStrings' {
  const strings: IFaqWebPartStrings;
  export = strings;
}
