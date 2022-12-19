import React, { useEffect } from "react";
import { SpreadsheetComponent, SheetsDirective } from "@syncfusion/ej2-react-spreadsheet";
import { IframeAction, IframeMode } from "../../interface/enum";
import { isObject } from "@syncfusion/ej2-base";

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

  const iframeActions = async (event: MessageEvent) => {
    const { action, key, value } = event.data;

    console.log("message from parent recieved:", action);

    switch (action) {
      case IframeMode.PREVIEW:
      case IframeMode.TEMPORARY_PREVIEW:
        spreadsheet.allowEditing = false;
        spreadsheet.allowCellFormatting = false;
        spreadsheet.showRibbon = false;
        break;
      case IframeMode.EDIT:
        spreadsheet.allowEditing = true;
        spreadsheet.allowCellFormatting = true;
        spreadsheet.showRibbon = true;
        break;
      case IframeAction.SAVE:
      case IframeAction.FREE_DRAFT:
        event.source!.postMessage(
          {
            action,
            key,
            value: JSON.stringify(await spreadsheet.saveAsJson()),
          },
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          // @ts-ignore
          "*",
        );
        localStorage.setItem(key, JSON.stringify(await spreadsheet.saveAsJson()));
        break;
      case IframeAction.LOAD:
        if (isObject(JSON.parse(value))) {
          spreadsheet.open(JSON.parse(value));
        }
        break;
      default:
        break;
    }
  };

  useEffect(() => {
    window.addEventListener("message", iframeActions, false);

    return () => {
      window.removeEventListener("message", iframeActions);
    };
  }, []);

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
    </div>
  );
};
export default ExcelComponent;
