import {WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDcrChangeRequestsProps {
  description: string;
  itemsPerPage: number;
  siteurl: string;
  context: WebPartContext;
  prioritySelectedKey: string;
  spWebUrl: string;
}
