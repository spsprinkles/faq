declare interface IFaqWebPartStrings {
  ListNameFieldDescription: string;
  ListNameFieldLabel: string;
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
