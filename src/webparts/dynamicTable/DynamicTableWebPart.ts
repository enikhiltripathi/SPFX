import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "DynamicTableWebPartStrings";
import DynamicTable from "./components/DynamicTable";
import {
  IDynamicTableProps,
  ISpFxRichTextEditorProps,
} from "./components/IDynamicTableProps";

export interface IDynamicTableWebPartProps {
  description: string;
}

export default class DynamicTableWebPart extends BaseClientSideWebPart<ISpFxRichTextEditorProps> {
  public render(): void {
    const element: React.ReactElement<ISpFxRichTextEditorProps> = React.createElement(
      DynamicTable,
      {
        renderTable: "",
        tableData: "",
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
