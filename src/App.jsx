import React, { useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import { DndProvider, useDrag, useDrop } from "react-dnd";
import { HTML5Backend } from "react-dnd-html5-backend";
import "./app.css";

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

        marginBottom: "4px",
        backgroundColor: "white",
        cursor: "move",
      }}
    >
      <table width={"100%"} style={{ textAlign: "right" }}>
        <tr>
          <th></th>
          <th>שם החשבון</th>
          <th>מס' ע.מורשה</th>
        </tr>
        <tr>
          <td>{index + 1}</td>
          <td>{row["שם החשבון"]}</td>
          <td>{row["מס' ע.מורשה"]}</td>
        </tr>
      </table>
    </div>
  );
};

// רכיב droppable לשמות בקובץ השני
const DroppableRow = ({ row, index, onDrop, ind }) => {
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
        marginBottom: "4px",
        backgroundColor: isOver ? "lightgreen" : "white",
      }}
    >
      <table width={"100%"} style={{ textAlign: "right" }}>
        <tr>
          <th></th>
          <th>שם</th>
          <th>עוסק מורשה</th>
        </tr>
        <tr>
          <td>{ind + 1}</td>
          <td>{row["שם"]}</td>
          <td>{row["עוסק מורשה"]}</td>
        </tr>
      </table>
    </div>
  );
};

function App() {
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);
  const [loadMargUp, setLoadMargUp] = useState(false);
  const [manualMatches, setManualMatches] = useState([]);
  const [sorted1, setSorted1] = useState([]);
  const [sorted2, setSorted2] = useState([]);
  const [unmatchedFromFile1, setUnmatchedFromFile1] = useState([]);
  const [allDataFromFile2, setAllDataFromFile2] = useState([]);
  const [searchTerm, setSearchTerm] = useState(""); // משתנה לחיפוש
  // פונקציה לטיפול בחיפוש
  const handleSearch = (e) => {
    setSearchTerm(e.target.value);
  };

  // פונקציה לסינון נתונים לפי החיפוש
  const filterData = (data, columns) => {
    return data.filter((row) => columns.some((col) => row[col]?.toString().includes(searchTerm)));
  };

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

  const cleanString = (str) => {
    return str
      .replaceAll(" ", "")
      .replaceAll("\r", "")
      .replaceAll("\t", "")
      .replaceAll('"', "")
      .replaceAll("-", "")
      .replaceAll(".", "")
      .toLowerCase();
  };

  // פונקציה להתאמת שמות דומה
  const matchBySimilarNames = (data1, data2, key1, key2) => {
    const similarMatches = [];

    data1.forEach((row1) => {
      const cleanedKey1 = cleanString(row1[key1]);

      const matchedRow = data2.find((row2) => {
        const cleanedKey2 = cleanString(row2[key2]);
        return cleanedKey1.includes(cleanedKey2) || cleanedKey2.includes(cleanedKey1);
      });

      if (matchedRow) {
        similarMatches.push({ row1, matchedRow });
      }
    });

    return similarMatches;
  };

  // פונקציה לטיפול בשחרור (drop)
  const handleDrop = (row1, row2, index1) => {
    setManualMatches((prev) => [...prev, { ...row2, "מפתח חשבון": row1["מפתח חשבון"] }]);

    // הסרת השורה מהרשימה הראשונה של הנהלת חשבונות לפי שם החשבון
    setSorted1((prev) => prev.filter((r) => r["שם החשבון"] !== row1["שם החשבון"]));

    // הסרת השורה מהרשימה השנייה של זיו לפי שם
    setSorted2((prev) => prev.filter((r) => r["שם"] !== row2["שם"]));
  };

  const handleFileUpload = async () => {
    if (!file1 || !file2) {
      alert("Please upload both files.");
      return;
    }
    setLoadMargUp(true);
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
      const similarNameMatches = matchBySimilarNames(
        sortByHebrewAlphabet(unmatchedData1, "שם החשבון"),
        sortByHebrewAlphabet(data2, "שם"),
        "שם החשבון",
        "שם"
      );

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
  const handleExport = () => {
    // יצירת הנתונים הממוזגים (קובץ זיו עם עמודת "מפתח חשבון")
    const updatedDataFromFile2 = allDataFromFile2.map((row) => {
      const match = manualMatches.find((match) => match["שם"] === row["שם"]);
      return {
        ...row,
        "מפתח חשבון": match ? match["מפתח חשבון"] : "",
      };
    });

    // נתונים שלא נמצאה להם התאמה מהקובץ הראשון
    const unmatchedDataFromFile1 = unmatchedFromFile1;

    // יצירת Workbook חדש
    const workbook = XLSX.utils.book_new();

    // יצירת גליון ראשון לנתונים הממוזגים
    const worksheet1 = XLSX.utils.json_to_sheet(updatedDataFromFile2);
    console.log("🚀 ~ handleExport ~ worksheet1:", worksheet1);
    worksheet1["!dir"] = "rtl"; // הגדרת כיוון הגיליון מימין לשמאל

    XLSX.utils.book_append_sheet(workbook, worksheet1, "נתונים ממוזגים");

    // יצירת גליון שני לשורות מהקובץ הראשון ללא התאמה
    const worksheet2 = XLSX.utils.json_to_sheet(unmatchedDataFromFile1);
    worksheet2["!dir"] = "rtl"; // הגדרת כיוון הגיליון מימין לשמאל

    // יישור התאים בגליון מימין לשמאל
    for (let cell in worksheet2) {
      if (cell[0] !== "!") {
        worksheet2[cell].s = { alignment: { readingOrder: 2, horizontal: "right" } }; // יישור לימין
      }
    }

    XLSX.utils.book_append_sheet(workbook, worksheet2, "ללא התאמה מהקובץ הנהלת חשבונות");

    // כתיבת הקובץ והורדתו
    const xlsxData = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([xlsxData], { type: "application/octet-stream" });
    saveAs(blob, "הנהלת חשבונות.xlsx");
  };

  return (
    <DndProvider backend={HTML5Backend}>
      <div style={{ display: "flex", flexDirection: "column", gap: "15px", alignItems: "center" }}>
        <h1>מיזוג עמודת "מפתח חשבון" של קובץ הנהלת חשבונות עם קובץ זיו</h1>

        <form action="#" style={{ display: !loadMargUp ? "block" : "none" }}>
          <div className="input-file-container">
            <input
              className="input-file"
              id="my-file"
              type="file"
              accept=".xlsx"
              onChange={(e) => {
                document.querySelector(".file-return-0").innerHTML = e.target.value;
                setFile1(e.target.files[0]);
              }}
            />
            <label tabIndex="0" htmlFor="my-file" className="input-file-trigger">
              העלה קובץ הנהלת חשבונות
            </label>
          </div>
          <p className="file-return-0"></p>
        </form>

        <form action="#" style={{ display: !loadMargUp ? "block" : "none" }}>
          <div className="input-file-container">
            <input
              className="input-file"
              id="my-file"
              type="file"
              accept=".xlsx"
              onChange={(e) => {
                console.log("🚀 ~ App ~ e:", e);
                document.querySelector(".file-return-1").innerHTML = e.target.value;

                setFile2(e.target.files[0]);
              }}
            />
            <label tabIndex="0" htmlFor="my-file" className="input-file-trigger">
              העלה קובץ זיו
            </label>
          </div>
          <p className="file-return-1"></p>
        </form>

        <button
          className="btn"
          onClick={handleFileUpload}
          style={{ display: !loadMargUp ? "block" : "none" }}
        >
          טען קבצים והתחל במיזוג האוטומטי
        </button>

        <h2
          style={{
            display: !loadMargUp ? "none" : "flex",
          }}
        >
          גרור/י טבלה מתאימה מעמודת הנהלת חשבונות לעמודת זיו
        </h2>

        {/* שדה החיפוש */}
        <input
          type="text"
          placeholder="חפש לפי שם, מס' ע.מורשה או עוסק מורשה"
          value={searchTerm}
          onChange={handleSearch}
          style={{
            display: loadMargUp ? "block" : "none",
            padding: "10px",
            marginBottom: "20px",
            width: "300px",
          }}
        />
        <button
          className="btn"
          onClick={handleExport}
          style={{
            display: !loadMargUp ? "none" : "flex",
          }}
        >
          יצא את הקובץ הממוזג
        </button>
        {/* הצגת התוצאות לאחר חיפוש */}
        <div
          style={{
            display: loadMargUp ? "flex" : "none",
            justifyContent: "space-between",
            gap: "20px",
          }}
        >
          <div>
            <h3>שמות בקובץ הנהלת חשבונות</h3>
            {filterData(sorted1, ["שם החשבון", "מס' ע.מורשה"]).map((row1, index) => (
              <DraggableRow key={index} row={row1} index={index} />
            ))}
          </div>

          <div>
            <h3>שמות בקובץ זיו</h3>
            {filterData(sorted2, ["שם", "עוסק מורשה"]).map((row2, index) => (
              <DroppableRow key={index} row={row2} index={index} onDrop={handleDrop} ind={index} />
            ))}
          </div>
        </div>
      </div>
    </DndProvider>
  );
}

export default App;
