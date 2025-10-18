import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICorporateDirectoryProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  documentLibrary: string;
  csvFile: string;
}
