import React, { useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import { DndProvider, useDrag, useDrop } from "react-dnd";
import { HTML5Backend } from "react-dnd-html5-backend";

// פונקציה לבדוק אם יש לפחות מילה אחת תואמת בין שני שמות
const hasMatchingWord = (str1, str2) => {
  const words1 = str1.split(" ");
  const words2 = str2.split(" ");
  return words1.some((word) => words2.includes(word));
};

// רכיב draggable לשמות מתוך הקובץ הראשון
const DraggableRow = ({ row, index }) => {
  const [{ isDragging }, drag] = useDrag({
    type: "row",
    item: { row, index },
    collect: (monitor) => ({
      isDragging: !!monitor.isDragging(),
    }),
  });

  return (
    <div
      ref={drag}
      style={{
        opacity: isDragging ? 0.5 : 1,
        padding: "8px",
        border: "1px solid gray",
        marginBottom: "4px",
        backgroundColor: "white",
        cursor: "move",
      }}
    >
      {row["שם החשבון"]}
    </div>
  );
};

// רכיב droppable לשמות בקובץ השני
const DroppableRow = ({ row, index, onDrop }) => {
  const [{ isOver }, drop] = useDrop({
    accept: "row",
    drop: (item) => {
      if (hasMatchingWord(item.row["שם החשבון"], row["שם"])) {
        onDrop(item.row, row, item.index);
      }
    },
    collect: (monitor) => ({
      isOver: !!monitor.isOver(),
    }),
  });

  return (
    <div
      ref={drop}
      style={{
        padding: "8px",
        border: "1px solid gray",
        marginBottom: "4px",
        backgroundColor: isOver ? "lightgreen" : "white",
      }}
    >
      {row["שם"]}
    </div>
  );
};

function App() {
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);
  const [manualMatches, setManualMatches] = useState([]);
  const [sorted1, setSorted1] = useState([]);
  const [sorted2, setSorted2] = useState([]);
  const [unmatchedFromFile1, setUnmatchedFromFile1] = useState([]);
  const [allDataFromFile2, setAllDataFromFile2] = useState([]);

  // פונקציה לקריאת קובץ XLSX
  const readXLSXFile = async (file) => {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    return worksheet;
  };

  // פונקציה למיון לפי א-ב
  const sortByHebrewAlphabet = (data, key) => {
    return data.sort((a, b) => a[key].localeCompare(b[key], "he"));
  };

  // פונקציה למיזוג אוטומטי לפי ID
  const autoMatchById = (data1, data2, idKey1, idKey2) => {
    const matchedRows = [];
    const unmatchedData1 = [];

    data1.forEach((row1) => {
      const matchedRow = data2.find((row2) => row2[idKey2] === row1[idKey1]);
      if (matchedRow) {
        matchedRows.push({ ...matchedRow, "מפתח חשבון": row1["מפתח חשבון"] });
      } else {
        unmatchedData1.push(row1);
      }
    });

    return { matchedRows, unmatchedData1 };
  };

  // פונקציה להתאמת שמות דומה
  const matchBySimilarNames = (data1, data2, key1, key2) => {
    const similarMatches = [];

    data1.forEach((row1) => {
      const matchedRow = data2.find(
        (row2) => row1[key1].includes(row2[key2]) || row2[key2].includes(row1[key1])
      );
      if (matchedRow) {
        similarMatches.push({ ...matchedRow, "מפתח חשבון": row1["מפתח חשבון"] });
      }
    });

    return similarMatches;
  };

  // פונקציה לטיפול בשחרור (drop)
  const handleDrop = (row1, row2, index1) => {
    setManualMatches((prev) => [...prev, { ...row2, "מפתח חשבון": row1["מפתח חשבון"] }]);
    setSorted1((prev) => prev.filter((_, i) => i !== index1)); // הסרה מהרשימה הראשונה
    setSorted2((prev) => prev.filter((r) => r !== row2)); // הסרה מהרשימה השנייה
  };

  const handleFileUpload = async () => {
    if (!file1 || !file2) {
      alert("Please upload both files.");
      return;
    }

    try {
      const data1 = await readXLSXFile(file1);
      const data2 = await readXLSXFile(file2);

      // מיזוג אוטומטי לפי "מס' ע.מורשה" ו-"עוסק מורשה"
      const { matchedRows, unmatchedData1 } = autoMatchById(
        data1,
        data2,
        "מס' ע.מורשה",
        "עוסק מורשה"
      );

      // חיפוש התאמות לפי שמות "שם החשבון" ו-"שם"
      const similarNameMatches = matchBySimilarNames(unmatchedData1, data2, "שם החשבון", "שם");

      // הצגת הרשימה למשתמש להתאמה ידנית עם drag-and-drop
      setSorted1(
        sortByHebrewAlphabet(
          unmatchedData1.filter((row) => !similarNameMatches.includes(row)),
          "שם החשבון"
        )
      );
      setSorted2(sortByHebrewAlphabet(data2, "שם"));

      // שמירה על כל הנתונים מהקובץ השני
      setAllDataFromFile2(data2);

      // שמירה על ההתאמות האוטומטיות
      setManualMatches((prev) => [
        ...prev,
        ...matchedRows, // התאמות לפי מספר זיהוי
        ...similarNameMatches, // התאמות לפי שמות דומים
      ]);

      // שמירה על שורות שלא נמצאו להם התאמות מהקובץ הראשון
      setUnmatchedFromFile1(unmatchedData1);
    } catch (error) {
      console.error("Error processing files:", error);
      alert("An error occurred while processing the files.");
    }
  };

  // פונקציה ליצוא קובץ XLSX
  const exportXLSX = (data, fileName) => {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, fileName);
    const xlsxData = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([xlsxData], { type: "application/octet-stream" });
    saveAs(blob, `${fileName}.xlsx`);
  };

  const handleExport = () => {
    // יצוא קובץ עם עמודת "מפתח חשבון" בקובץ השני
    const updatedDataFromFile2 = allDataFromFile2.map((row) => {
      const match = manualMatches.find((match) => match["שם"] === row["שם"]);
      return {
        ...row,
        "מפתח חשבון": match ? match["מפתח חשבון"] : "",
      };
    });

    // כלל כל שורות מהקובץ הראשון שלא נמצאה להן התאמה
    const finalData = [...updatedDataFromFile2, ...unmatchedFromFile1];

    exportXLSX(finalData, "הנהלת חשבונות");
  };

  return (
    <DndProvider backend={HTML5Backend}>
      <div style={{ display: "flex", flexDirection: "column", gap: "15px" }}>
        <h1>Merge XLSX Files with Drag-and-Drop Matching</h1>
        <input type="file" accept=".xlsx" onChange={(e) => setFile1(e.target.files[0])} />
        <input type="file" accept=".xlsx" onChange={(e) => setFile2(e.target.files[0])} />
        <button onClick={handleFileUpload}>Load Files for Matching</button>

        <div style={{ display: "flex", justifyContent: "space-between", gap: "20px" }}>
          <div>
            <h3>שמות בקובץ הראשון</h3>
            {sorted1.map((row1, index) => (
              <DraggableRow key={index} row={row1} index={index} />
            ))}
          </div>

          <div>
            <h3>שמות בקובץ השני</h3>
            {sorted2.map((row2, index) => (
              <DroppableRow key={index} row={row2} index={index} onDrop={handleDrop} />
            ))}
          </div>
        </div>

        <button onClick={handleExport}>Export Matched Data</button>
      </div>
    </DndProvider>
  );
}

export default App;
