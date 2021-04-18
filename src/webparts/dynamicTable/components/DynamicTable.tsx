import * as React from "react";
import styles from "./DynamicTable.module.scss";
import {
  IDynamicTableProps,
  ISpFxRichTextEditorProps,
  ISpFxRichTextEditorState,
} from "./IDynamicTableProps";
import { escape } from "@microsoft/sp-lodash-subset";
import ReactTable from "./ReactTable";
import {
  SPHttpClient,
  SPHttpClientResponse,
  SPHttpClientConfiguration,
} from "@microsoft/sp-http";
import { SPListOperations } from "spfxhelper";
import { Log, ServiceScope } from "@microsoft/sp-core-library";

export default class DynamicTable extends React.Component<
  ISpFxRichTextEditorProps,
  ISpFxRichTextEditorState
> {
  constructor(
    props: ISpFxRichTextEditorProps,
    state: ISpFxRichTextEditorState
  ) {
    super(props);
    this.state = {
      ReactTableResult: [],
      HTTPClient: "",
      SPSiteURL: "",
    };
  }
  handleTableData = (tableRowColl) => {
    this.setState({ ReactTableResult: tableRowColl });
  };

  public render(): React.ReactElement<IDynamicTableProps> {
    return (
      <ReactTable
        tableData={this.handleTableData}
        renderTable={this.state.ReactTableResult}
        context={this.props.context}
        siteUrl={this.props.context.pageContext.web.absoluteUrl}
      ></ReactTable>
    );
  }
}
