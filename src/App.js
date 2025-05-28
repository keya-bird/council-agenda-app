/* global XLSX */
import React, { useState, useEffect } from 'react';

// Main App component for the dynamic table
function App() {
  // State to hold the table data.
  // It tries to load data from localStorage on initial render, otherwise starts with an empty array.
  const [data, setData] = useState(() => {
    try {
      const savedData = localStorage.getItem('dynamicTableData');
      return savedData ? JSON.parse(savedData) : [];
    } catch (error) {
      console.error("Failed to parse data from localStorage:", error);
      return []; // Return empty array if parsing fails
    }
  });

  // State for the input fields when adding a new row
  const [newRow, setNewRow] = useState({ time: '', department: '', issue: '', presenter: '' });

  // State to keep track of which row is currently being edited
  const [editingRowId, setEditingRowId] = useState(null);

  // State for the input fields when editing an existing row
  const [editedRow, setEditedRow] = useState({ time: '', department: '', issue: '', presenter: '' });

  // State for displaying messages to the user (e.g., success/error for file upload)
  const [message, setMessage] = useState('');

  // State to track if the XLSX library has been successfully loaded
  const [isXLSXLoaded, setIsXLSXLoaded] = useState(false);

  // State to track the ID of the currently highlighted row
  const [highlightedRowId, setHighlightedRowId] = useState(null);

  // Helper function to convert Excel serial time (decimal) to HH:MM string
  const excelSerialToTime = (serial) => {
    if (typeof serial !== 'number' || isNaN(serial)) {
      return String(serial); // Return as is if not a valid number
    }

    // Excel serial time is a fraction of a day
    const totalSeconds = serial * 24 * 60 * 60;
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);

    // Format hours and minutes to ensure two digits (e.g., 09 instead of 9)
    const formattedHours = String(hours).padStart(2, '0');
    const formattedMinutes = String(minutes).padStart(2, '0');

    return `${formattedHours}:${formattedMinutes}`;
  };

  // Effect to dynamically load the XLSX library from CDN
  useEffect(() => {
    // Check if XLSX is already available globally to prevent re-loading
    if (typeof XLSX !== 'undefined') {
      setIsXLSXLoaded(true);
      return;
    }

    const scriptId = 'xlsx-cdn-script'; // Unique ID for the script tag
    // Check if the script tag already exists in the document head
    if (document.getElementById(scriptId)) {
      // If it exists, assume it's either loading or loaded.
      // We'll rely on the 'isXLSXLoaded' state to confirm readiness.
      return;
    }

    const script = document.createElement('script');
    script.id = scriptId;
    script.src = "https://unpkg.com/xlsx/dist/xlsx.full.min.js";
    script.async = true; // Load script asynchronously

    // Set onload handler for successful script loading
    script.onload = () => {
      setIsXLSXLoaded(true);
      setMessage('XLSX library loaded successfully.');
    };

    // Set onerror handler for failed script loading
    script.onerror = () => {
      setMessage('Error: Failed to load XLSX library from CDN. Please check your network connection.');
      setIsXLSXLoaded(false);
    };

    // Append the script to the document's head
    document.head.appendChild(script);

    // Cleanup function: remove the script when the component unmounts
    return () => {
      const existingScript = document.getElementById(scriptId);
      if (existingScript) {
        document.head.removeChild(existingScript);
      }
    };
  }, []); // Empty dependency array ensures this effect runs only once on mount

  // Effect to save data to localStorage whenever the 'data' state changes
  useEffect(() => {
    try {
      localStorage.setItem('dynamicTableData', JSON.stringify(data));
    } catch (error) {
      console.error("Failed to save data to localStorage:", error);
    }
  }, [data]); // Dependency array: this effect runs whenever 'data' changes

  // Handles changes in the input fields for adding a new row
  const handleNewRowChange = (e) => {
    const { name, value } = e.target;
    setNewRow(prev => ({ ...prev, [name]: value }));
  };

  // Handles changes in the input fields for editing an existing row
  const handleEditedRowChange = (e) => {
    const { name, value } = e.target;
    setEditedRow(prev => ({ ...prev, [name]: value }));
  };

  // Adds a new row to the table
  const addRow = () => {
    // Check if all fields for the new row are filled
    if (newRow.time && newRow.department && newRow.issue && newRow.presenter) {
      // Create a unique ID for the new row using a timestamp
      const newId = Date.now();
      // Add the new row to the data state
      setData(prevData => [...prevData, { id: newId, ...newRow }]);
      // Clear the input fields for adding a new row
      setNewRow({ time: '', department: '', issue: '', presenter: '' });
      setMessage(''); // Clear any previous messages
    } else {
      // Set a message if any field is empty
      setMessage("Please fill all fields to add a new row.");
    }
  };

  // Starts the editing process for a specific row
  const startEdit = (row) => {
    setEditingRowId(row.id); // Set the ID of the row being edited
    // Populate the editedRow state with the current row's data
    setEditedRow({ time: row.time, department: row.department, issue: row.issue, presenter: row.presenter });
    setHighlightedRowId(null); // Unhighlight any row when starting edit
  };

  // Saves the changes made to an edited row
  const saveEdit = (id) => {
    setData(prevData =>
      prevData.map(row =>
        // If the row ID matches, update the row with the edited data; otherwise, keep the original row
        row.id === id ? { ...row, ...editedRow } : row
      )
    );
    setEditingRowId(null); // Exit editing mode
    setEditedRow({ time: '', department: '', issue: '', presenter: '' }); // Clear edited row state
  };

  // Cancels the editing process
  const cancelEdit = () => {
    setEditingRowId(null); // Exit editing mode
    setEditedRow({ time: '', department: '', issue: '', presenter: '' }); // Clear edited row state
  };

  // Deletes a row from the table with confirmation
  const deleteRow = (id) => {
    // Use the global confirm dialog for confirmation
    const isConfirmed = window.confirm("Are you sure you want to delete this row?");
    
    if (isConfirmed) {
      // Filter out the row with the matching ID to delete it
      setData(prevData => prevData.filter(row => row.id !== id));
      if (highlightedRowId === id) { // If the deleted row was highlighted, unhighlight it
        setHighlightedRowId(null);
      }
    }
  };

  // Handles clicking on a table row to highlight it
  const handleRowClick = (id) => {
    // If the clicked row is already highlighted, unhighlight it. Otherwise, highlight it.
    setHighlightedRowId(prevId => (prevId === id ? null : id));
  };

  // Handles the Excel file upload
  const handleFileUpload = (event) => {
    // Prevent file upload if XLSX library is not loaded
    if (!isXLSXLoaded) {
      setMessage('XLSX library is still loading or failed to load. Please wait a moment and try again.');
      return;
    }

    const file = event.target.files[0];
    if (file) {
      // Check if the file type is an Excel spreadsheet
      if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        setMessage('Error: Please upload a valid Excel file (.xlsx or .xls).');
        return;
      }

      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames[0]; // Get the first sheet
          const worksheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // Get data as array of arrays

          if (json.length === 0) {
            setMessage('Error: The uploaded Excel file is empty or could not be parsed.');
            return;
          }

          // Assuming the first row contains headers: Time, Department, Issue, Presenter
          const headers = json[0].map(header => String(header).toLowerCase().trim());
          const expectedHeaders = ['time', 'department', 'issue', 'presenter'];

          // Basic validation for headers
          const hasAllExpectedHeaders = expectedHeaders.every(header => headers.includes(header));

          if (!hasAllExpectedHeaders) {
            setMessage('Error: Excel file headers do not match expected: Time, Department, Issue, Presenter. Please check your file.');
            return;
          }

          const newRows = json.slice(1).map((rowArray) => {
            const newRowData = {};
            expectedHeaders.forEach((expectedHeader) => {
              const excelHeaderIndex = headers.indexOf(expectedHeader);
              let cellValue = '';
              if (excelHeaderIndex !== -1 && rowArray[excelHeaderIndex] !== undefined) {
                cellValue = rowArray[excelHeaderIndex];
              }

              // Special handling for 'time' column
              if (expectedHeader === 'time') {
                newRowData[expectedHeader] = excelSerialToTime(cellValue);
              } else {
                newRowData[expectedHeader] = String(cellValue);
              }
            });
            return { id: Date.now() + Math.random(), ...newRowData }; // Add unique ID
          }).filter(row => expectedHeaders.some(h => row[h] && row[h].trim() !== '')); // Filter out completely empty rows

          if (newRows.length > 0) {
            setData(prevData => [...prevData, ...newRows]); // Append new rows
            setMessage(`Successfully uploaded ${newRows.length} rows from Excel.`);
          } else {
            setMessage('No valid data rows found in the Excel file after parsing.');
          }

        } catch (error) {
          console.error("Error reading Excel file:", error);
          setMessage(`Error: Failed to read Excel file: ${error.message}`);
        }
      };

      reader.readAsArrayBuffer(file); // Read file as ArrayBuffer for XLSX
    } else {
      setMessage('No file selected.');
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-4 sm:p-8 font-sans antialiased">
      <div className="max-w-4xl mx-auto bg-white rounded-xl shadow-lg p-6 sm:p-8">
        <h1 className="text-3xl sm:text-4xl font-extrabold text-gray-800 mb-8 text-center">
          Council Agenda Table
        </h1>

        {/* Data Table Section - Moved up */}
        <div className="overflow-x-auto bg-white rounded-lg shadow-md mb-8"> {/* Added mb-8 for spacing */}
          <table className="min-w-full divide-y divide-gray-200">
            <thead className="bg-gray-50">
              <tr>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Time</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Department</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Issue</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Presenter</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Actions</th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {data.length === 0 ? (
                <tr>
                  <td colSpan="5" className="px-6 py-4 whitespace-nowrap text-sm text-gray-500 text-center">
                    No data available. Add some entries or upload an Excel file!
                  </td>
                </tr>
              ) : (
                data.map((row) => (
                  <tr
                    key={row.id}
                    onClick={() => handleRowClick(row.id)}
                    className={`cursor-pointer transition duration-150 ease-in-out ${
                      highlightedRowId === row.id ? 'bg-yellow-200' : 'hover:bg-gray-50'
                    }`}
                  >
                    {editingRowId === row.id ? (
                      <>
                        <td className="px-6 py-4 whitespace-nowrap">
                          <input
                            type="text"
                            name="time"
                            placeholder="HH:MM"
                            value={editedRow.time}
                            onChange={handleEditedRowChange}
                            className="p-2 border border-gray-300 rounded-md w-full focus:outline-none focus:ring-1 focus:ring-indigo-400"
                          />
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap">
                          <input
                            type="text"
                            name="department"
                            value={editedRow.department}
                            onChange={handleEditedRowChange}
                            className="p-2 border border-gray-300 rounded-md w-full focus:outline-none focus:ring-1 focus:ring-indigo-400"
                          />
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap">
                          <input
                            type="text"
                            name="issue"
                            value={editedRow.issue}
                            onChange={handleEditedRowChange}
                            className="p-2 border border-gray-300 rounded-md w-full focus:outline-none focus:ring-1 focus:ring-indigo-400"
                          />
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap">
                          <input
                            type="text"
                            name="presenter"
                            value={editedRow.presenter}
                            onChange={handleEditedRowChange}
                            className="p-2 border border-gray-300 rounded-md w-full focus:outline-none focus:ring-1 focus:ring-indigo-400"
                          />
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
                          <div className="flex space-x-2">
                            <button
                              onClick={(e) => { e.stopPropagation(); saveEdit(row.id); }} // Stop propagation to prevent row highlight
                              className="text-green-600 hover:text-green-900 font-semibold transition duration-150 ease-in-out transform hover:scale-105"
                            >
                              Save
                            </button>
                            <button
                              onClick={(e) => { e.stopPropagation(); cancelEdit(); }} // Stop propagation to prevent row highlight
                              className="text-gray-500 hover:text-gray-700 font-semibold transition duration-150 ease-in-out transform hover:scale-105"
                            >
                              Cancel
                            </button>
                          </div>
                        </td>
                      </>
                    ) : (
                      <>
                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">{row.time}</td>
                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">{row.department}</td>
                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">{row.issue}</td>
                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">{row.presenter}</td>
                        <td className="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
                          <div className="flex space-x-2">
                            <button
                              onClick={(e) => { e.stopPropagation(); startEdit(row); }} // Stop propagation to prevent row highlight
                              className="text-indigo-600 hover:text-indigo-900 font-semibold transition duration-150 ease-in-out transform hover:scale-105"
                            >
                              Edit
                            </button>
                            <button
                              onClick={(e) => { e.stopPropagation(); deleteRow(row.id); }} // Stop propagation to prevent row highlight
                              className="text-red-600 hover:text-red-900 font-semibold transition duration-150 ease-in-out transform hover:scale-105"
                            >
                              Delete
                            </button>
                          </div>
                        </td>
                      </>
                    )}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {/* Add New Row Section - Moved down */}
        <div className="mb-8 p-6 bg-blue-50 rounded-lg shadow-inner">
          <h2 className="text-2xl font-bold text-blue-800 mb-4">Add New Entry</h2>
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-4">
            <input
              type="text"
              name="time"
              placeholder="Time (HH:MM)"
              value={newRow.time}
              onChange={handleNewRowChange}
              className="p-3 border border-blue-200 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-400 transition duration-200"
            />
            <input
              type="text"
              name="department"
              placeholder="Department"
              value={newRow.department}
              onChange={handleNewRowChange}
              className="p-3 border border-blue-200 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-400 transition duration-200"
            />
            <input
              type="text"
              name="issue"
              placeholder="Issue"
              value={newRow.issue}
              onChange={handleNewRowChange}
              className="p-3 border border-blue-200 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-400 transition duration-200"
            />
            <input
              type="text"
              name="presenter"
              placeholder="Presenter"
              value={newRow.presenter}
              onChange={handleNewRowChange}
              className="p-3 border border-blue-200 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-400 transition duration-200"
            />
          </div>
          <button
            onClick={addRow}
            className="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3 px-6 rounded-lg shadow-md transition duration-300 ease-in-out transform hover:scale-105 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2"
          >
            Add Row
          </button>
        </div>

        {/* Upload Excel Section - Moved down */}
        <div className="mb-8 p-6 bg-green-50 rounded-lg shadow-inner">
          <h2 className="text-2xl font-bold text-green-800 mb-4">Upload Excel Document</h2>
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFileUpload}
            className="block w-full text-sm text-gray-700
                       file:mr-4 file:py-2 file:px-4
                       file:rounded-full file:border-0
                       file:text-sm file:font-semibold
                       file:bg-green-50 file:text-green-700
                       hover:file:bg-green-100 mb-4"
            disabled={!isXLSXLoaded} /* Disable input until XLSX is loaded */
          />
          {!isXLSXLoaded && (
            <p className="mt-2 text-sm text-center font-medium text-yellow-700">
              Loading XLSX library...
            </p>
          )}
          {message && (
            <p className="mt-2 text-sm text-center font-medium"
               style={{ color: message.startsWith('Error') ? 'red' : (message.startsWith('Successfully') ? 'green' : 'inherit') }}>
              {message}
            </p>
          )}
        </div>
      </div>
    </div>
  );
}

export default App;