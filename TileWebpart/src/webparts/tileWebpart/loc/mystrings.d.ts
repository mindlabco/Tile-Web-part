declare interface ITileWebpartWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'TileWebpartWebPartStrings' {
  const strings: ITileWebpartWebPartStrings;
  export = strings;
}

interface JQuery
{
  rateYo(options: IOptions | string): Function;
}
interface IOptions {
  rating?: number;  
  starWidth?:string; 
  readOnly?:boolean;
  ratedFill?:string;
}