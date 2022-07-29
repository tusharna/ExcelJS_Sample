import React, { Component } from "react";
import DataGrid, {
  Column,
  Export,
  GroupPanel
} from "devextreme-react/data-grid";
import Button from "devextreme-react/button";
import ExcelJS from "exceljs/dist/es5/exceljs.browser";
import saveAs from "file-saver";
import service from "./data.js";

class App extends Component {
  constructor(props) {
    super(props);
    this.employees = service.getEmployees();

    this.state = {
      excelFilterEnabled: true
    };
  }

  handleExportClick = e => {
    this.excelExport(this.state.instanceDataGrid);
  };

  excelExport = DataGrid => {
    var ExcelJSWorkbook = new ExcelJS.Workbook();
    var worksheet = ExcelJSWorkbook.addWorksheet("ExcelJS sheet");
    var columns = DataGrid.getVisibleColumns();

    worksheet.mergeCells("A2:I2");

    const customCell = worksheet.getCell("A2");
    customCell.font = {
      name: "Comic Sans MS",
      family: 4,
      size: 20,
      underline: true,
      bold: true
    };

    customCell.value = "Custom header here";

    var headerRow = worksheet.addRow();
    worksheet.getRow(4).font = { bold: true };

    for (let i = 0; i < columns.length; i++) {
      let currentColumnWidth = DataGrid.option().columns[i].width;
      worksheet.getColumn(i + 1).width =
        currentColumnWidth !== undefined ? currentColumnWidth / 6 : 20;
      let cell = headerRow.getCell(i + 1);
      cell.value = columns[i].caption;
    }

    if (this.state.excelFilterEnabled === true) {
      worksheet.autoFilter = {
        from: {
          row: 3,
          column: 1
        },
        to: {
          row: 3,
          column: columns.length
        }
      };
    }

    // eslint-disable-next-line no-unused-expressions
    this.state.excelFilterEnabled === true
      ? (worksheet.views = [{ state: "frozen", ySplit: 3 }])
      : undefined;

    worksheet.properties.outlineProperties = {
      summaryBelow: false,
      summaryRight: false
    };

    DataGrid.getController("data")
      .loadAll()
      .then(function(allItems) {
        for (let i = 0; i < allItems.length; i++) {
          var dataRow = worksheet.addRow();
          if (allItems[i].rowType === "data") {
            dataRow.outlineLevel = 1;
          }
          for (let j = 0; j < allItems[i].values.length; j++) {
            let cell = dataRow.getCell(j + 1);
            cell.value = allItems[i].values[j];
          }
        }

        const rowCount = worksheet.rowCount;
        worksheet.mergeCells(`A${rowCount}:I${rowCount + 1}`);
        worksheet.getRow(1).font = { bold: true };
        worksheet.getCell(`A${rowCount}`).font = {
          name: "Comic Sans MS",
          family: 4,
          size: 20,
          underline: true,
          bold: true
        };

        worksheet.getCell(`A${rowCount}`).value = "Custom Footer here";

        ExcelJSWorkbook.xlsx.writeBuffer().then(function(buffer) {
          saveAs(
            new Blob([buffer], { type: "application/octet-stream" }),
            `${DataGrid.option().export.fileName}.xlsx`
          );
        });
      });
  };

  render() {
    return (
      <div>
        <DataGrid
          id={"gridContainer"}
          dataSource={this.employees}
          showBorders={true}
          showColumnHeaders={true}
          onCellPrepared={this.onCellPrepared}
          onContentReady={this.onContentReady}
        >
          <Column dataField={"Prefix"} caption={"Title"} width={60} />
          <Column dataField={"FirstName"} />
          <Column dataField={"LastName"} />
          <Column dataField={"City"} />
          <Column dataField={"State"} />

          <Column dataField={"Position"} width={130} />
          <Column dataField={"BirthDate"} dateType={"date"} width={130} />
          <Column dataField={"HireDate"} dateType={"date"} width={100} />
          <Column
            dataField={"SaleAmount"}
            alighment={"right"}
            format={"currency"}
          />
          <Export
            enabled={true}
            fileName="Employees"
            excelFilterEnabled={true}
            customizeExcelCell={this.customizeExcelCell}
          />
          <GroupPanel visible={true} />
        </DataGrid>
        <Button text="export" type="danger" onClick={this.handleExportClick} />
      </div>
    );
  }

  onContentReady = e => {
    var instanceGrid = e.component.instance();

    this.setState({
      excelFilterEnabled: instanceGrid.option().export.excelFilterEnabled,
      instanceDataGrid: instanceGrid
    });
  };

  onCellPrepared(e) {
    if (e.rowType === "data") {
      if (e.data.OrderDate < new Date(2014, 2, 3)) {
        e.cellElement.classList.add("oldOrder");
      }
      if (e.data.SaleAmount > 15000) {
        if (e.column.dataField === "Employee") {
          e.cellElement.classList.add("highAmountOrder_employee");
        }
        if (e.column.dataField === "SaleAmount") {
          e.cellElement.classList.add("highAmountOrder_saleAmount");
        }
      }
    }
  }

  customizeExcelCell(options) {
    if (options.gridCell.rowType === "data") {
      if (options.gridCell.data.SaleAmount > 15000) {
        if (options.gridCell.column.dataField === "Employee") {
          options.font.bold = true;
        }
        if (options.gridCell.column.dataField === "SaleAmount") {
          options.backgroundColor = "#FFBB00";
          options.font.color = "#000000";
        }
      }
    }
  }
}

export default App;
