declare interface IBannerWebPartStrings {
  BannerConfigName: string;
  BannerTextField: string;
  BannerSecondaryText: string;
  BannerImageUrlField: string;
  BannerLinkField: string;
  BannerNumberField: string;
  BannerParallaxField: string;
  BannerValidationNotImage: string;
  BannerPlaceholderIconText: string;
  BannerPlaceholderDescription: string;
  BannerPlaceholderBtnLabel: string;
  

  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'BannerWebPartStrings' {
  const strings: IBannerWebPartStrings;
  export = strings;
}
