/* import react and office ui fabric */
import { Icon } from "office-ui-fabric-react";
import * as React from "react";
import styles from "./DynamicTable.module.scss";
import { ISpFxRichTextEditorProps } from "./IDynamicTableProps";
import { Log } from "@microsoft/sp-core-library";
import { SPListOperations } from "spfxhelper";
import {
  SPHttpClient,
  SPHttpClientResponse,
  SPHttpClientConfiguration,
} from "@microsoft/sp-http";

/* Initial loading of table column properties */
const tableColProps = [
  {
    id: 1,
    NounName: "",
    CurrentPN: "",
    NewPart: "",
    Option: "",
    Supplier: "",
    EJR: "",
  },
];

/* Declare React Table State variables */
export interface IReactTableState {
  tableHeaders: any;
  rows: any;
  rowLimit: number;
  addButtonText: string;
  removeButtonText: string;
  HTTPClient: any;
  SPSiteURL: any;
}

export default class ReactTable extends React.Component<
  ISpFxRichTextEditorProps,
  IReactTableState
> {
  /* React Table constructor */
  constructor(props: ISpFxRichTextEditorProps) {
    super(props);
    this.state = {
      tableHeaders: [
        "Invoice Number",
        "Item code",
        "Item Desc",
        "Base Price",
        "Invoice Date",
        "Title",
      ],
      rows: this.props.renderTable,
      rowLimit: 20,
      addButtonText: "Add New Row",
      removeButtonText: "Delete Row",
      HTTPClient: "",
      SPSiteURL: "",
    };
  }

  /* This event will fire on cell change */
  handleChange = (index) => (evt) => {
    try {
      var item = {
        id: evt.target.id,
        name: evt.target.name,
        value: evt.target.value,
      };
      var rowsArray = this.state.rows;
      var newRow = rowsArray.map((row, i) => {
        for (var key in row) {
          if (key == item.name && row.id == item.id) {
            row[key] = item.value;
          }
        }
        return row;
      });
      this.setState({ rows: newRow });
      this.props.tableData(rowsArray);
    } catch (error) {
      console.log("Error in React Table handle change : " + error);
    }
  };

  /* This event will fire on adding new row */
  handleAddRow = () => {
    try {
      var id = Math.floor(Math.random() * 300);
      const tableColProps = {
        id: id,
        InvoiceNumber: "",
        ItemCode: "",
        ItemDesc: "",
        ItemBasePrice: "",
        InvoiceDate: "",
        Title: "",
      };
      if (this.state.rows.length < this.state.rowLimit) {
        this.state.rows.push(tableColProps);
        this.setState(this.state.rows);
      } else {
        alert("Add row limit exceeds");
      }
    } catch (error) {
      console.log("Error in React Table handle Add Row : " + error);
    }
  };

  /* This event will fire on remove row */
  handleRemoveRow = () => {
    try {
      var rowsArray = this.state.rows;
      if (rowsArray.length > 1) {
        var newRow = rowsArray.slice(0, -1);
        this.setState({ rows: newRow });
      }
      this.props.tableData(newRow);
    } catch (error) {
      console.log("Error in React Table handle Remove Row : " + error);
    }
  };

  /* This event will fire on remove specific row */
  handleRemoveSpecificRow = (idx) => () => {
    try {
      const rows = [this.state.rows];
      if (rows.length > 1) {
        rows.splice(idx, 1);
      }

      this.setState({ rows });
    } catch (error) {
      console.log("Error in React Table handle Remove Specific Row : " + error);
    }
  };

  /* This event will fire on next properties update */

  componentWillReceiveProps(nextProps) {
    try {
      if (nextProps.renderTable.length > 0) {
        this.setState({ rows: nextProps.renderTable });
      } else {
        this.setState({ rows: tableColProps });
      }
    } catch (error) {
      console.log(
        "Error in React Table component will receive props : " + error
      );
    }
  }

  render() {
    let list = this.state.rows.map((item, idx) => {
      return (
        <tr key={idx}>
          <td>
            <input
              type="text"
              name="InvoiceNumber"
              value={this.state.rows[idx].InvoiceNumber}
              onChange={this.handleChange(idx)}
              id={this.state.rows[idx].id}
            />
          </td>

          <td>
            <input
              type="text"
              name="ItemCode"
              value={this.state.rows[idx].ItemCode}
              onChange={this.handleChange(idx)}
              id={this.state.rows[idx].id}
            />
          </td>

          <td>
            <input
              type="text"
              name="ItemDesc"
              value={this.state.rows[idx].ItemDesc}
              onChange={this.handleChange(idx)}
              id={this.state.rows[idx].id}
            />
          </td>

          <td>
            <input
              type="text"
              name="ItemBasePrice"
              value={this.state.rows[idx].ItemBasePrice}
              onChange={this.handleChange(idx)}
              id={this.state.rows[idx].id}
            />
          </td>

          <td>
            <input
              type="text"
              name="InvoiceDate"
              value={this.state.rows[idx].InvoiceDate}
              onChange={this.handleChange(idx)}
              id={this.state.rows[idx].id}
            />
          </td>

          <td>
            <input
              type="text"
              name="Title"
              value={this.state.rows[idx].Title}
              onChange={this.handleChange(idx)}
              id={this.state.rows[idx].id}
            />
          </td>
        </tr>
      );
    });

    return (
      <div className={styles.mainDynTable + "container"}>
        <table id={styles.dynVPITable}>
          <thead>
            <tr>
              {this.state.tableHeaders.map(function (headerText) {
                return <th> {headerText} </th>;
              })}
            </tr>
          </thead>

          <tbody>{list}</tbody>
        </table>

        <div className={styles["add-remove-icons"]}>
          <span
            id="add-row"
            onClick={this.handleAddRow}
            className={styles["document-icons-area"]}
          >
            <Icon iconName="CalculatorAddition" className="ms-IconExample" />
          </span>
          <span
            id="delete-row"
            onClick={this.handleRemoveRow}
            className={styles["document-icons-area"]}
          >
            <Icon iconName="CalculatorSubtract" className="ms-IconExample" />
          </span>
        </div>

        <button type="button" onClick={() => this.AddItems("He")}>
          Add to SharePoint
        </button>
      </div>
    );
  }

  private get oListOperation(): SPListOperations {
    return SPListOperations.getInstance(
      this.state.HTTPClient as SPHttpClient,
      this.state.SPSiteURL,
      "WeBpart"
    );
  }

  /* Add the items to the list */
  private AddItems(workordertype: string): void {
    try {
      var dynamicTableData = "";
      var workOrderList = "Sales Items";
      var resTable = this.ValidateDynamicTable(this.state.rows);
      if (resTable.length > 0) {
        dynamicTableData = JSON.stringify({ resTable });
        dynamicTableData = dynamicTableData.replace(/[{}]/g, "");
      }

      const bodyforadding: string = JSON.stringify({
        __metadata: { type: "SP.Data.Sales_x0020_ItemsListItem" },
        TableContent: dynamicTableData,
      });

      /*Add items to sharepoint list */
      this.AddListItems(
        this.props.context.spHttpClient,
        this.props.siteUrl,
        workOrderList,
        bodyforadding
      );
    } catch (error) {
      console.log("Error in addItems : " + error);
    }
  }

  /*Validate Dynamic table */
  public ValidateDynamicTable(vpiReactTableResult) {
    try {
      var stateReactArr;
      if (vpiReactTableResult == "" || vpiReactTableResult == null) {
        stateReactArr = [];
      } else {
        if (vpiReactTableResult.length > 0) {
          stateReactArr = vpiReactTableResult;
        } else {
          stateReactArr = [];
        }
      }
      return stateReactArr;
    } catch (error) {
      console.log("Error in ValidateDynamicTable : " + error);
    }
  }

  /** Add item to sharepoint list */
  public AddListItems(spHttpClientVPI, siteURL, sharepointList, body): void {
    try {
      spHttpClientVPI
        .post(
          siteURL +
            "/_api/web/lists/getbytitle('" +
            sharepointList +
            "')/items",
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "odata-version": "",
            },
            body: body,
          }
        )
        .then((): void => {
          alert("Item Added Successfully!!!");
        });
    } catch (error) {
      console.log("Error in AddListItems : " + error);
    }
  }
}
