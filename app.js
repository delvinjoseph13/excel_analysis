

let file1Data = null, file2Data = null;


document.getElementById('fileInput1').addEventListener('change', (event) => {
    handleFileSelect1(event);
}, false);

document.getElementById('fileInput2').addEventListener('change', (event) => {
    handleFileSelect2(event);
}, false);

function handleFileSelect1(event) {
    resetAppState()
    const file = event.target.files[0];
    if (!file) return;

    // Clear previous file2Data when a new gogst file is selected
    file2Data = null;
    const inputElement2 = document.getElementById('fileInput2');
    if (inputElement2) {
        inputElement2.value = ''; // Clear the B2B file input
    }
    

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetNames = workbook.SheetNames;

        const section = document.querySelector(".file-section");

        // Remove existing dropdown if any
        const existingDropdown = document.getElementById('sheetDropdown1');
        if (existingDropdown) existingDropdown.parentElement.remove();

        if (sheetNames.length > 1) {
            const dropdown = document.createElement('select');
            dropdown.id = 'sheetDropdown1';
            dropdown.innerHTML = sheetNames.map(name => `<option value="${name}">${name}</option>`).join('');

            const confirmButton = document.createElement('button');
            confirmButton.textContent = 'Select';
            confirmButton.addEventListener('click', () => {
                const selectedSheet = dropdown.value;
                const worksheet = workbook.Sheets[selectedSheet];
                file1Data = XLSX.utils.sheet_to_json(worksheet);

                renderFileList(file, 'fileInput1', 'file1');
                // compareData();
                dropdown.parentElement.remove();
            });

            const dropdownContainer = document.createElement('div');
            dropdownContainer.appendChild(dropdown);
            dropdownContainer.appendChild(confirmButton);

            section.appendChild(dropdownContainer);
        } else {




            const worksheet = workbook.Sheets[sheetNames[0]];

           const range = XLSX.utils.decode_range(worksheet['!ref']);
            range.s.r = 5; // Set start row to 5 (zero-based index)
            worksheet['!ref'] = XLSX.utils.encode_range(range);

            file1Data = XLSX.utils.sheet_to_json(worksheet,{ header: 1 });
  
                        // console.log(file1Data)

            file1Data=file1Data.map((row)=>({
                "Zone_Name":row[2],
                "Zone_Area_Name":row[4],
                "TM.5 - ltr":row[5],
                "HM1LTR - ltr":row[6],
                "CWM500 - ltr":row[7],
                "HM525 - ltr":row[8],
                "S500 - ltr":row[9],
                "S1000 - ltr":row[10],
                "HFM1 - ltr":row[11],
                "CM750 - ltr":row[12],
                "FZRMILK - ltr":row[13]
            }))

           const dropdown=document.createElement("select");
            dropdown.classList.add("dropdown");
            dropdown.id='sheetDropdown5';
            const uniqueZone = [
  ...new Set(
    file1Data
      .map(zone => zone.Zone_Name)
      .filter(zoneName => zoneName && zoneName !== "Zone Name")
  )
];

// dropdown.innerHTML = uniqueZone.map(zoneName => {`
//     <option value="" disabled>select Zone Name</option>
//     <option value="${zoneName}">${zoneName}</option>
//     `})
// .join('');

dropdown.innerHTML =
  `<option value="" disabled selected>Select Zone Name</option>` +
  uniqueZone
    .map(zoneName => `<option value="${zoneName}">${zoneName}</option>`)
    .join('');


const dropdrownForZoneArea=document.createElement("select");
dropdrownForZoneArea.classList.add("dropdown");
dropdrownForZoneArea.id='sheetDropDown6';

//   const uniqueZoneArea=[...new Set(file1Data.map(zonearea=>zonearea.Zone_Area_Name).filter(zonearea=>zonearea && zonearea!=="Zone Area Name" && zonearea.Zone_Name===dropdown.value ))]

// const uniqueZoneArea=[...new Set(file1Data.filter(zonearea=>zonearea.Zone_Name===dropdown.value).map(zonearea=>zonearea.Zone_Area_Name)
//  ) ]  

//   dropdrownForZoneArea.innerHTML=uniqueZoneArea.map(zonearea=>`<option value="${zonearea}">${zonearea}</option>`)
updateFilterZoneArea(dropdown.value,dropdrownForZoneArea)
dropdown.addEventListener("change", () => {
    updateFilterZoneArea(dropdown.value, dropdrownForZoneArea);
});

  const dropdownForSelectProduct=document.createElement("select")  
  dropdownForSelectProduct.classList.add("dropdown");
  dropdownForSelectProduct.id='dropdownproduct'

  const uniqueProducts=["TM.5 - ltr","HM1LTR - ltr","CWM500 - ltr","HM525 - ltr","S500 - ltr",
    "S1000 - ltr","HFM1 - ltr","CM750 - ltr","FZRMILK - ltr"]
  dropdownForSelectProduct.innerHTML=
  `<option value="" disabled selected>Select Product</option>` +
  uniqueProducts.map(product=>`<option value="${product}">${product}</option>`)
  
  console.log(file1Data)

  const confirmButtonFilter=document.createElement('button');
  confirmButtonFilter.classList.add("Download-button")
  confirmButtonFilter.textContent='Select';
  confirmButtonFilter.addEventListener('click',()=>{
    fetchDataFilter(dropdown,dropdrownForZoneArea,dropdownForSelectProduct)
  })

  const clearButton=document.createElement('button');
  clearButton.classList.add("clear-button")
  clearButton.textContent='Clear';
  clearButton.addEventListener('click',clearAllData)

    //   clearButton.addEventListener('click', () => {
    //   file1Data = null;
    //    file2Data = null;
    //     fileList.removeChild(listItem);
    // });

  

            renderFileList(file, 'fileInput1', 'file1');
            // compareData();

            const dropdownContainer = document.createElement('div');
            dropdownContainer.classList.add("dropdown-container");
            dropdownContainer.appendChild(dropdown);
            dropdownContainer.appendChild(dropdrownForZoneArea)
            dropdownContainer.appendChild(dropdownForSelectProduct)
            dropdownContainer.appendChild(confirmButtonFilter)
            dropdownContainer.appendChild(clearButton)
               const sheetSelection=document.querySelector("#zone-selection")
            sheetSelection.appendChild(dropdownContainer);
        }
    };

    reader.readAsBinaryString(file);
}

function clearAllData() {
    // Reset data
    file1Data = [];
    file2Data = [];

    // Clear tables
    document.getElementById("table-container").innerHTML = "";

    // Clear dropdowns
    document.getElementById("zone-selection").innerHTML = "";

    // Clear file list
    document.getElementById("file-list").innerHTML = "";

    // Reset file inputs
    document.getElementById("fileInput1").value = "";
    document.getElementById("fileInput2").value = "";
}




function updateFilterZoneArea(selectedZone,zoneAreaDropDown){
    const uniqueZoneArea=[...new Set(file1Data.filter(zonearea=>zonearea.Zone_Name===selectedZone).map(zonearea=>zonearea.Zone_Area_Name))]

    zoneAreaDropDown.innerHTML=
     `<option value="" disabled selected>Select Area</option>` +
    uniqueZoneArea.map(area=>`<option value="${area}">${area}</option>`).join('')
 

}

function fetchDataFilter(zoneSelect, zoneAreaSelect, productSelect) {

    const selectedZone =document.getElementById("sheetDropdown5").value;
    console.log(selectedZone)

    const zoneAreaOption=document.getElementById("sheetDropDown6");
     
    const allOptionValues=Array.from(zoneAreaOption.options).map(option=>option.value).filter(value=>value);

    

    const selectedZoneArea = zoneAreaSelect.value;
    const selectedProduct = productSelect.value;

    const filteredData = file1Data.filter(
        row =>
            row.Zone_Name === selectedZone &&
            row.Zone_Area_Name === selectedZoneArea
    );

    console.log("filtered data",filteredData)

    if (!filteredData) {
        alert("No data found");
        return;
    }

    console.log(filteredData[0][selectedProduct])

    const tableContainer = document.getElementById("table-container");
    tableContainer.innerHTML = ""; // ✅ CLEAR ONCE HERE
     

     const productTotal=filteredData.reduce((acc,curr)=>{
        return acc+(Number(curr[selectedProduct]) || 0)
     },0)


    const tableData = {
        "Zone Name": filteredData[0].Zone_Name,
        "Zone Area": filteredData[0].Zone_Area_Name,
        "Product Name": selectedProduct,
        "Product Value":productTotal ?? 0
    };

    const fetchTotalItemValue = file1Data
        .filter(row => row.Zone_Name === selectedZone)
        .reduce((total, row) => total + Number(row[selectedProduct] || 0), 0);
    
    const filterreduce=file1Data.filter(row=>row.Zone_Name ===selectedZone)
    
    console.log(selectedProduct)
    const chartValue=filterreduce.map(val=>val[selectedProduct])

    console.log(chartValue)

    console.log(filterreduce)

    const dataforTotal = {
        "Zone Name": selectedZone,
        "Product Name": selectedProduct,
        "Product Total": fetchTotalItemValue
    };

    console.log("datafrottoal",dataforTotal)

const filterProductList = Object.keys(file1Data[0])
  .filter(key => key !== "Zone_Name" && key !== "Zone_Area_Name");


//   console.log("poro",filterProductList)

const totalForEachProduct = filterProductList.map(product => {
  const total = file1Data.reduce((sum, row) => {
    let value = row[product];

    // Convert safely
    value = typeof value === "string"
      ? value.replace(/,/g, "").trim()
      : value;

    return sum + (Number(value) || 0);
  }, 0);

  return { product, total };
});

// console.log(totalForEachProduct);

const resultObject = totalForEachProduct.reduce((acc, item) => {
  acc[item.product] = item.total;
  return acc;
}, {});

// console.log("resultObject",resultObject);
    const zoneTotals=file1Data.reduce((acc,row)=>{
        const zone=row.Zone_Name;
        const value=Number(row[selectedProduct] || 0);

        if(!zone || Number.isNaN(value)) return acc;

        acc[zone]=(acc[zone] || 0) +value
        return acc
    },{})

    const zoneLabel=Object.keys(zoneTotals);
    const zoneValue=Object.values(zoneTotals)
    
    // console.log(zoneTotals)

    addHeadingText("Total Product");
    renderTable(resultObject, "table-container");

    // ✅ First table
    addHeadingText("Zone Wise Report");
    renderTable(tableData, "table-container");

    // ✅ Second table
    addHeadingText("Total Report");
    renderTable(dataforTotal, "table-container");
     
    
       
createCanvas(
  allOptionValues,
  chartValue,
  "Zone Area Wise Product",
  "areaChart"
);

createCanvas(
  zoneLabel,
  zoneValue,
  "Zone Wise Product Total",
  "zoneChart"
);


    


}


function resetAppState() {
    file1Data = [];
    file2Data = [];

    document.getElementById("table-container").innerHTML = "";
    document.getElementById("zone-selection").innerHTML = "";
    document.getElementById("file-list").innerHTML = "";
}



function handleFileSelect2(event) {
    const file = event.target.files[0];
    if (!file) return;


    const reader = new FileReader();
    reader.onload = function (e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetNames = workbook.SheetNames;

        const section = document.querySelector(".file-section");

        // Remove existing dropdown if any
        const existingDropdown = document.getElementById('sheetDropdown2');
        if (existingDropdown) existingDropdown.parentElement.remove();

        if (sheetNames.length > 1) {
            const dropdown = document.createElement('select');
            dropdown.classList.add("dropdown")
            dropdown.id = 'sheetDropdown2';
            dropdown.innerHTML = sheetNames.map(name => `<option value="${name}">${name}</option>`).join('');

            const confirmButton = document.createElement('button');
            confirmButton.classList.add("Download-button")
            confirmButton.textContent = 'Select';
            confirmButton.addEventListener('click', () => {
                const selectedSheet = dropdown.value;
                const worksheet = workbook.Sheets[selectedSheet];

                // Skip the first 4 rows
                const range = XLSX.utils.decode_range(worksheet['!ref']);
                range.s.r = 4; // Set start row to 5 (zero-based index)
                worksheet['!ref'] = XLSX.utils.encode_range(range);

                file2Data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                // Extract specific columns
                file2Data = file2Data.map(row => ({
                    'GSTIN of supplier': row[0],
                    'Trade/Legal name': row[1],
                    'Invoice number': row[2],
                    'Invoice Date': row[4]
                }));

                // Console the extracted data
                console.log("Extracted B2B Data:", file2Data);

                renderFileList(file, 'fileInput2', 'file2');
                // compareData();
                dropdown.parentElement.remove();
            });

            const dropdownContainer = document.createElement('div');
            dropdownContainer.classList.add("dropdown-container");
            dropdownContainer.appendChild(dropdown);
            dropdownContainer.appendChild(confirmButton);
            const sheetSelection=document.querySelector("#sheet-selection")
            sheetSelection.appendChild(dropdownContainer);
            // section.appendChild(dropdownContainer);

        } else {
            const worksheet = workbook.Sheets[sheetNames[0]];

            // Skip the first 4 rows
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            range.s.r = 5; // Set start row to 5 (zero-based index)
            worksheet['!ref'] = XLSX.utils.encode_range(range);

            file2Data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Extract specific columns
            file2Data = file2Data.map(row => ({
                 "Zone_Name":row[2],
                "Zone_Area_Name":row[4],
                "TM.5 - ltr":row[5],
                "HM1LTR - ltr":row[6],
                "CWM500 - ltr":row[7],
                "HM525 - ltr":row[8],
                "S500 - ltr":row[9],
                "S1000 - ltr":row[10],
                "HFM1 - ltr":row[11],
                "CM750 - ltr":row[12],
                "FZRMILK - ltr":row[13]
            }));

            file1Data =[...file1Data,...file2Data]


            // Console the extracted data
            console.log("Extracted B2B Data:", file1Data);

            renderFileList(file, 'fileInput2', 'file2');
            // compareData();
        }
    };

    reader.readAsBinaryString(file);
}

const fileList = document.getElementById('file-list');

function renderFileList(file, fileKey) {
    const existingItem = fileList.querySelector(`[data-file="${fileKey}"]`);
    // if (existingItem) fileList.removeChild(existingItem);

    const listItem = document.createElement('li');
    listItem.setAttribute('data-file', fileKey);
    listItem.textContent = `${file.name}`;
    listItem.classList.add('file-item');

    const deleteButton = document.createElement('button');
    deleteButton.innerHTML = '<i class="material-icons md-48">delete</i>';
    deleteButton.classList.add('delete-btn');

    deleteButton.addEventListener('click', () => {
        if (fileKey === 'file1') file1Data = null;
        if (fileKey === 'file2') file2Data = null;
        fileList.removeChild(listItem);
    });

    listItem.appendChild(deleteButton);
    fileList.appendChild(listItem);
}

function addHeadingText(text){
    const headingText=document.createElement("h2");
    headingText.classList.add("heading-text")
    headingText.textContent=text;
    document.getElementById('table-container').appendChild(headingText);
}



function renderTable(data, containerId, clear = false) {
    const tableContainer = document.getElementById(containerId);

    if (clear) {
        tableContainer.innerHTML = '';
    }

    const table = document.createElement('table');
    table.classList.add('data-table');

    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    Object.keys(data).forEach(key => {
        const th = document.createElement('th');
        th.textContent = key;
        headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    const row = document.createElement('tr');

    Object.values(data).forEach(value => {
        const td = document.createElement('td');
        td.textContent = value;
        row.appendChild(td);
    });

    tbody.appendChild(row);
    table.appendChild(tbody);

    tableContainer.appendChild(table);
}




function generateExcelWithColors(data) {
    if (typeof ExcelJS === 'undefined') {
        console.error("ExcelJS is not loaded. Make sure you have included the correct script or installed it properly.");
        return;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Updated Data");

    // Define the correct column order explicitly
    const expectedColumns = [
        "SL", "REF.ID", "HSN CODE", "Head of Account", "GST NO.", "NAME OF PARTY",
        "STATE", "INV.NO.", "DATE", "VALUE OF GOODS", "GST Rate%", "CGST", "SGST", 
        "IGST", "TOTAL", "Purchase/Service", "Matching Status"
    ];

    // Add header row with explicit column order
    worksheet.addRow(expectedColumns).font = { bold: true };

    // Add data rows
    data.forEach(item => {
        // Ensure columns are in the correct order
        const rowValues = expectedColumns.map(header => item[header] || "");
        const row = worksheet.addRow(rowValues);

        // Apply red color if "Matching Status" is "Matching"
        if (item["Matching Status"] === "Matching") {
            row.eachCell(cell => {
                cell.fill = {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: { argb: "489838" } 
                };
                cell.font = { color: { argb: "FFFFFFFF" } }; // White text
            });
        }
    });

    // Generate and download the Excel file
    workbook.xlsx.writeBuffer().then(buffer => {
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const fileInput1 = document.getElementById('fileInput1');
        const fileName = fileInput1.files[0].name;
        const outputFileName = `Updated_${fileName}`;

        saveAs(blob, outputFileName);
    }).catch(error => console.error("Error generating Excel file:", error));
}


function fetchDetails(){

}

let myChart = null;
const charts=[];

// function compareData() {
//     if (file1Data && file2Data) {
//         // Define the expected column order
//         const expectedColumns = [
//             "SL", "REF.ID", "HSN CODE", "Head of Account", "GST NO.", "NAME OF PARTY",
//             "STATE", "INV.NO.", "DATE", "VALUE OF GOODS", "GST Rate%", "CGST", "SGST", 
//             "IGST", "TOTAL", "Purchase/Service"
//         ];

//         file1Data = file1Data.map((item) => {
//             const isMatching = file2Data.some((b2bItem) => {
//                 return (
//                     item['GST NO.'] === b2bItem['GSTIN of supplier'] ||
//                     item['INV.NO.'] === b2bItem['Invoice number'] ||
//                     item['DATE'] === b2bItem['Invoice Date'] ||
//                     item['NAME OF PARTY'] === b2bItem['Trade/Legal name']
//                 );
//             });

//             // Ensure all expected columns exist
//             let rowData = {};
//             expectedColumns.forEach(col => {
//                 rowData[col] = item[col] || "";  // Fill missing values with empty string
//             });

//             // Append "Matching Status" at the end
//             rowData["Matching Status"] = isMatching ? "Matching" : "Non-Matching";

//             return rowData;
//         });

//         console.log("Updated File 1 Data with Matching Status:", file1Data);

//         renderTable(file1Data);
//         generateExcelWithColors(file1Data);
//     }
// }



// const ctx=document.getElementById("chart");

function createCanvas(labels, values, title, chartId) {
  const chartContainer = document.getElementById("chart-container");
  
  chartContainer.style.marginTop="20px"

  // Wrapper
  const wrapper = document.createElement("div");
  wrapper.className = "chart-box";


  // Heading
  const heading = document.createElement("h2");
  heading.textContent = title;

  // Canvas (UNIQUE ID)
  const canvas = document.createElement("canvas");
  canvas.id = chartId;



  wrapper.appendChild(heading);
  wrapper.appendChild(canvas);
  chartContainer.appendChild(wrapper);

  // Create chart
  const chart = new Chart(canvas, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: title,
        data: values,
        borderWidth: 1
      }]
    },
    options: {
      responsive: true,
      scales: {
        y: { beginAtZero: true }
      }
    }
  });

  charts.push(chart); // store reference
}







