import React from "react";
import { SpreadsheetComponent, SheetsDirective } from "@syncfusion/ej2-react-spreadsheet";

const ExcelComponent = () => {
  let spreadsheet: SpreadsheetComponent;

  function onCreated(): void {
    spreadsheet.cellFormat(
      {
        fontWeight: "bold",
        textAlign: "center",
        verticalAlign: "middle",
      },
      "A1:F1",
    );
    spreadsheet.numberFormat("$#,##0.00", "F2:F31");
  }

  const handleSave = async () => {
    const data = await spreadsheet.saveAsJson();
    localStorage.setItem("data", JSON.stringify(data));
    console.log(data);
  };

  const handleLoad = () => {
    const data = localStorage.getItem("data");
    spreadsheet.open(JSON.parse(data as string));
  };

  return (
    <div className="control-pane">
      <div className="control-section spreadsheet-control">
        <SpreadsheetComponent
          openUrl="https://ej2services.syncfusion.com/production/web-services/api/spreadsheet/open"
          saveUrl="https://ej2services.syncfusion.com/production/web-services/api/spreadsheet/save"
          ref={(ssObj) => {
            // eslint-disable-next-line @typescript-eslint/ban-ts-comment
            // @ts-ignore
            spreadsheet = ssObj;
          }}
          // eslint-disable-next-line react/jsx-no-bind
          created={onCreated.bind(this)}
        >
          <SheetsDirective />
        </SpreadsheetComponent>
      </div>
      <button onClick={handleSave}>save</button>
      <button onClick={handleLoad}>load</button>
    </div>
  );
};
export default ExcelComponent;
