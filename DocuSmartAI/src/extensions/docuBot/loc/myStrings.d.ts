declare interface IDocuBotCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'DocuBotCommandSetStrings' {
  const strings: IDocuBotCommandSetStrings;
  export = strings;
}
