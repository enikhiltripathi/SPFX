import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDynamicTableProps {
  description: string;
}

export interface ISpFxRichTextEditorProps {
  tableData?: any;
  renderTable?: any;
  context: WebPartContext;
  siteUrl: any;
}

export interface ISpFxRichTextEditorState {
  ReactTableResult: any;
  HTTPClient: any;
  SPSiteURL: any;
}
