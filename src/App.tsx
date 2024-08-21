import reactLogo from "./assets/react.svg";
import viteLogo from "/vite.svg";
import "./App.css";
import * as XLSX from "xlsx";

function App() {
  const data = [
    {
      id: 1,
      name: "John",
      age: 25,
      tasks: [
        { date: "2024-08-01", noTasks: 6 },
        { date: "2024-08-02", noTasks: 7 },
      ],
    },
    {
      id: 2,
      name: "Jane",
      age: 30,
      tasks: [
        { date: "2024-08-01", noTasks: 6 },
        { date: "2024-08-02", noTasks: 7 },
      ],
    },
    {
      id: 3,
      name: "Bob",
      age: 35,
      tasks: [
        { date: "2024-08-01", noTasks: 6 },
        { date: "2024-08-02", noTasks: 7 },
      ],
    },
  ];

  const handleOnExport = async () => {
    // Flatten the data while keeping the structure
    const flattenedData = data.flatMap((person) =>
      person.tasks.map((task, index) => ({
        id: index === 0 ? person.id : "",
        name: index === 0 ? person.name : "",
        age: index === 0 ? person.age : "",
        taskDate: task.date,
        noTasks: task.noTasks,
      }))
    );

    const headers = [
      ["ID", "Name", "Age", "Tasks", "", ""],
      ["", "", "", "Date", "No of Tasks", ""],
    ];

    const worksheetData = [
      ...headers,
      ...flattenedData.map((row) => [row.id, row.name, row.age, row.taskDate, row.noTasks]),
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

    // Merge cells for the multi-level headers
    worksheet["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }, // Merge 'ID' column
      { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } }, // Merge 'Name' column
      { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } }, // Merge 'Age' column
      { s: { r: 0, c: 3 }, e: { r: 0, c: 4 } }, // Merge 'Tasks' header across 'Date' and 'No of Tasks'
    ];

    // Create a workbook and add the worksheet
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Tasks");

    XLSX.writeFile(workbook, `data.xlsx`);
  };

  return (
    <>
      <div>
        <a href="https://vitejs.dev" target="_blank">
          <img src={viteLogo} className="logo" alt="Vite logo" />
        </a>
        <a href="https://react.dev" target="_blank">
          <img src={reactLogo} className="logo react" alt="React logo" />
        </a>
      </div>
      <h1>Vite + React</h1>
      <div className="card">
        <button onClick={handleOnExport}>Export data</button>
        <p>
          Edit <code>src/App.tsx</code> and save to test HMR
        </p>
      </div>
      <p className="read-the-docs">Click on the Vite and React logos to learn more</p>
    </>
  );
}

export default App;
