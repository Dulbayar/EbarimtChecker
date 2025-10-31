/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    const runFillBtn = document.getElementById("run-fill");
    const createTableBtn = document.getElementById("create-table");

    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";
    if (runFillBtn) runFillBtn.onclick = runFill;
    if (createTableBtn) createTableBtn.onclick = () => tryCatch(createTable);
  }
});

export async function createTable() {
  try {
    await Excel.run(async (context) => {
      const tables = context.workbook.tables;
      tables.load("items");
      await context.sync();

      const tableExists = tables.items.some(t => t.name === "EbarimtTable");
      if (tableExists) {
        showMessage("Table 'EbarimtTable' already exists!");
        return;
      }

      const newWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const ebarimtTable = newWorksheet.tables.add("A1:D1", true);
      ebarimtTable.name = "EbarimtTable";
      ebarimtTable.getHeaderRowRange().values = [["РД", "Нэр", "Tin Дугаар", "msg"]];
      ebarimtTable.getRange().format.autofitColumns();
      ebarimtTable.getRange().format.autofitRows();

      await context.sync();
      showMessage("Table created successfully!");
    });
  } catch (error) {
    console.error(error);
    showMessage("Error creating table: " + error.message);
  }
}

export async function runFill() {
  try {
    await Excel.run(async (context) => {
      const tables = context.workbook.tables;
      tables.load("items");
      await context.sync();

      const tableExists = tables.items.some(t => t.name === "EbarimtTable");
      if (!tableExists) {
        showMessage("Please create the table first by clicking 'Талбар нээх'.");
        return;
      }

      const table = context.workbook.tables.getItem("EbarimtTable");
      const dataBodyRange = table.getDataBodyRange();
      dataBodyRange.load(["values", "rowCount"]);
      await context.sync();

      const rowCount = dataBodyRange.rowCount;

      for (let i = 0; i < rowCount; i++) {
        const regValue = dataBodyRange.values[i][0];
        if (!regValue) continue;

        const tinData = await fetchTinFromAPI(regValue);
        const nameData = await fetchNameFromAPI(tinData.tin);

        dataBodyRange.getCell(i, 1).values = [[nameData.name || ""]];
        const tinCell = dataBodyRange.getCell(i, 2);
        tinCell.values = [[tinData.tin || ""]];
        tinCell.numberFormat = [["0"]];
        dataBodyRange.getCell(i, 3).values = [[nameData.msg || tinData.msg || ""]];
      }

      await context.sync();
      table.getRange().format.autofitColumns();
      table.getRange().format.autofitRows();

      showMessage("Data filled successfully!");
    });
  } catch (error) {
    console.error("Error in runFill:", error);
    showMessage("Error: " + error.message);
  }
}

async function fetchTinFromAPI(id: string) {
  try {
    const url = `https://api.ebarimt.mn/api/info/check/getTinInfo?regNo=${id}`;
    const response = await fetch(url);
    const data = await response.json();
    return { tin: data.data || "", msg: data.msg || "" };
  } catch (error) {
    return { tin: "", msg: "Error fetching TIN" };
  }
}

async function fetchNameFromAPI(id: string) {
  try {
    const url = `https://api.ebarimt.mn/api/info/check/getInfo?tin=${id}`;
    const response = await fetch(url);
    const data = await response.json();
    return { name: data.data?.name || "", msg: data.msg || "" };
  } catch (error) {
    return { name: "", msg: "Error fetching name" };
  }
}

async function tryCatch(callback: () => Promise<void>) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

function showMessage(message: string) {
  const messageElement = document.getElementById("item-subject");
  if (messageElement) {
    messageElement.textContent = message;
    setTimeout(() => (messageElement.textContent = ""), 5000);
  }
}
