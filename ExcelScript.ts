function main(workbook: ExcelScript.Workbook) {
  console.log("Script started");

  // Get the tables
  let projectTable = workbook.getTable("Table1");
  let progressTable = workbook.getTable("Table11");
  console.log("Tables retrieved");

  // Get the data from the tables
  let projectTableData: (string | number | boolean)[][] = projectTable.getRange().getValues();
  let progressTableData: (string | number | boolean)[][] = progressTable.getRange().getValues();
  console.log("Data retrieved from tables");

  // Define status weights
  let statusWeights: { [key: string]: number } = {
    "Complete": 4,
    "UAT": 3,
    "Development": 2,
    "In Progress": 2,
    "NotApplicable": 0,
    "Scheduling": 1,
    "Scheduled": 1,
    "Blocked": 1,
    "Blocked-NoResponse": 1,
    "Blocked-Bandwidth": 1,
    "TBD": 1,
    "Blocked-NoRemoteAccess": 1,
    "Hold":1
  };

  // Extract PMs and their project counts from the project table
  let pmProjects = new Map<string, { count: number, score: number, targetScore: number }>();
  for (let i = 1; i < projectTableData.length; i++) {
    let pm = projectTableData[i][projectTable.getColumnByName("PM").getIndex()] as string;
    let status = projectTableData[i][projectTable.getColumnByName("Project Status").getIndex()] as string;

    if (pmProjects.has(pm)) {
      pmProjects.get(pm)!.count++;
      pmProjects.get(pm)!.score += statusWeights[status];
    } else {
      let initialScore = statusWeights[status];
      pmProjects.set(pm, { count: 1, score: initialScore, targetScore: 0 });
    }
  }

  // Calculate target score for each PM
  pmProjects.forEach((value, key) => {
    value.targetScore = value.count * 4;
  });

  console.log("PMs, project counts, scores, and target scores extracted:", Array.from(pmProjects.entries()));

  // Current date formatted as "DDMMMYY"
  let today = new Date();
  let day = String(today.getDate()).padStart(2, '0');
  let month = today.toLocaleString('default', { month: 'short' }).toUpperCase();
  let year = String(today.getFullYear()).slice(-2);
  let currentDate = day + month + year;
  console.log("Current date:", currentDate);

  // Score column name formatted as "Score-DDMMMYY"
  let scoreColumnName = "Score-" + currentDate;
  console.log("Score column name:", scoreColumnName);

  // Score% column name formatted as "Score%-DDMMMYY"
  let scorePercentageColumnName = "Score%-" + currentDate;
  console.log("Score% column name:", scorePercentageColumnName);

  // Function to add a new column if it doesn't exist
  function addColumnIfNeeded(table: ExcelScript.Table, columnName: string): number {
    let columns = table.getColumns();
    let columnIndex: number | undefined = undefined;

    for (let i = 0; i < columns.length; i++) {
      if (columns[i].getName() === columnName) {
        columnIndex = i;
        break;
      }
    }

    if (columnIndex === undefined) {
        table.addColumn();
        columns = table.getColumns(); // Refresh columns after adding
        columnIndex = columns.length - 1;  // New column is the last one
        table.getColumns()[columnIndex].setName(columnName);
        console.log("Added new column:", columnName, "at index:", columnIndex);
    }

    return columnIndex!;
  }

  // Add or retrieve Score column index
  let scoreColumnIndex = addColumnIfNeeded(progressTable, scoreColumnName);

  // Add or retrieve Target Score column index
  let targetScoreColumnIndex = addColumnIfNeeded(progressTable, "Target Score");

  // Add or retrieve Score% column index
  let scorePercentageColumnIndex = addColumnIfNeeded(progressTable, scorePercentageColumnName);

  // Add rows for each PM if they don't exist and set the project scores and percentages
  let pmEntries = Array.from(pmProjects.entries());

  for (let i = 0; i < pmEntries.length; i++) {
    let pm = pmEntries[i][0];
    let score = pmProjects.get(pm)!.score;
    let targetScore = pmProjects.get(pm)!.targetScore;

    let rowFound = false;

    for (let j = 1; j < progressTableData.length; j++) {
      if (progressTableData[j][0] === pm) {
        progressTable.getRangeBetweenHeaderAndTotal().getCell(j - 1, scoreColumnIndex).setValue(score);
        progressTable.getRangeBetweenHeaderAndTotal().getCell(j - 1, targetScoreColumnIndex).setValue(targetScore);
        progressTable.getRangeBetweenHeaderAndTotal().getCell(j - 1, scorePercentageColumnIndex).setValue((score / targetScore) * 100);
        rowFound = true;
        console.log("Updated existing row for PM:", pm, "with score:", score, "and target score:", targetScore);
        break;
      }
    }

    if (!rowFound) {
      let newRow: (string | number | boolean)[] = new Array(progressTableData[0].length).fill("");
      newRow[0] = pm;
      newRow[scoreColumnIndex] = score;
      newRow[targetScoreColumnIndex] = targetScore;
      newRow[scorePercentageColumnIndex] = (score / targetScore) * 100;
      progressTable.addRow(-1, newRow);
      console.log("Added new row for PM:", pm, "with score:", score, "and target score:", targetScore);
    }
  }

  console.log("Script ended");
}
