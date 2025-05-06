// Store data
let buildData = [];
let installCodes = [];
let addonData = [];
let selectedAddons = [];
let addonDescriptionsData = {}; // New object to store add-on descriptions

// Elements
const loadingElement = document.getElementById('loading');
const toolContainer = document.getElementById('tool-container');
const buildSearchInput = document.getElementById('build-search');
const codeSearchInput = document.getElementById('code-search');
const addonSearchInput = document.getElementById('addon-search');
const resultContainer = document.getElementById('result');
const selectedBuildNumber = document.getElementById('selected-build-number');
const buildDescriptionTextarea = document.getElementById('build-description');
const selectedAddonsList = document.getElementById('selected-addons-list');
const resetButton = document.getElementById('reset');
const printButton = document.getElementById('print');
const exportButton = document.getElementById('export-excel');

// Excel related elements
const excelUpload = document.getElementById('excel-upload');
const uploadStatus = document.getElementById('upload-status');

// Print template elements
const printContainer = document.getElementById('print-container');
const printBuildId = document.getElementById('print-build-id');
const printDate = document.getElementById('print-date');
const printNotes = document.getElementById('print-notes');
const printCodesTable = document.getElementById('print-codes-table').querySelector('tbody');
const printAddonsSection = document.getElementById('print-addons-section');
const printAddonsList = document.getElementById('print-addons-list');
const printTimestamp = document.getElementById('print-timestamp');
const closePrintButton = document.getElementById('close-print');
const actualPrintButton = document.getElementById('actual-print');

// Store custom descriptions for builds and add-ons
let customDescriptions = {};
let addonDescriptions = {};

// Initialize the app when the page loads
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

// Function to initialize the app
function initializeApp() {
    // Try to load saved custom descriptions from localStorage
    try {
        const savedBuildDescriptions = localStorage.getItem('buildCustomDescriptions');
        if (savedBuildDescriptions) {
            customDescriptions = JSON.parse(savedBuildDescriptions);
        }
        
        const savedAddonDescriptions = localStorage.getItem('addonCustomDescriptions');
        if (savedAddonDescriptions) {
            addonDescriptions = JSON.parse(savedAddonDescriptions);
        }
        
        // Try to load saved data from localStorage
        if (localStorage.getItem('buildData')) {
            loadDataFromLocalStorage();
        } else if (typeof configData !== 'undefined') {
            // Store the configuration data if no localStorage data exists
            loadDataFromConfig();
        } else {
            // No config data, show message to upload Excel
            uploadStatus.textContent = 'Please upload an Excel file with build and install code data.';
            uploadStatus.style.color = 'blue';
        }
    } catch (e) {
        console.error('Error loading saved data:', e);
        customDescriptions = {};
        addonDescriptions = {};
        
        // Fall back to config data if localStorage fails
        if (typeof configData !== 'undefined') {
            loadDataFromConfig();
        }
    }
    
    // Set up Excel file upload listener
    excelUpload.addEventListener('change', handleExcelUpload);
    
    // Set up Clear Saved Data button
    const clearSavedDataButton = document.getElementById('clear-saved-data');
    if (clearSavedDataButton) {
        clearSavedDataButton.addEventListener('click', clearSavedData);
    }
    
    // Hide loading and show tool
    loadingElement.style.display = 'none';
    toolContainer.style.display = 'block';
    
    // Set up event listeners for filters
    buildSearchInput.addEventListener('input', populateBuildsTable);
    codeSearchInput.addEventListener('input', filterInstallCodes);
    addonSearchInput.addEventListener('input', populateAddonsTable);
    
    // Save custom description when it changes
    buildDescriptionTextarea.addEventListener('input', function() {
        const selectedBuildNumberText = selectedBuildNumber.textContent;
        if (selectedBuildNumberText) {
            customDescriptions[selectedBuildNumberText] = buildDescriptionTextarea.value;
            // Save to localStorage
            try {
                localStorage.setItem('buildCustomDescriptions', JSON.stringify(customDescriptions));
            } catch (e) {
                console.error('Error saving descriptions:', e);
            }
        }
    });
    
    // Set up reset button
    resetButton.addEventListener('click', resetSelection);
    
    // Set up print button
    printButton.addEventListener('click', preparePrintView);
    
    // Set up export button
    if (exportButton) {
        exportButton.addEventListener('click', exportToExcel);
    }
    
    // Set up print view buttons
    closePrintButton.addEventListener('click', function() {
        printContainer.style.display = 'none';
    });
    
    actualPrintButton.addEventListener('click', function() {
        window.print();
    });
}

// Function to clear saved data
function clearSavedData() {
    if (confirm('Are you sure you want to clear all saved data? You will need to upload an Excel file again.')) {
        try {
            // Clear only the app data, keep user notes
            localStorage.removeItem('buildData');
            localStorage.removeItem('installCodes');
            localStorage.removeItem('addonData');
            localStorage.removeItem('addonDescriptionsData');
            localStorage.removeItem('lastUpdated');
            
            // Reload the page to start fresh
            window.location.reload();
        } catch (e) {
            console.error('Error clearing data:', e);
            alert('There was an error clearing the saved data.');
        }
    }
}

// Function to load data from localStorage
function loadDataFromLocalStorage() {
    try {
        buildData = JSON.parse(localStorage.getItem('buildData')) || [];
        installCodes = JSON.parse(localStorage.getItem('installCodes')) || [];
        addonData = JSON.parse(localStorage.getItem('addonData')) || [];
        addonDescriptionsData = JSON.parse(localStorage.getItem('addonDescriptionsData')) || {};
        
        const lastUpdated = localStorage.getItem('lastUpdated');
        if (lastUpdated) {
            uploadStatus.textContent = `Using saved data from: ${lastUpdated}`;
            uploadStatus.style.color = 'green';
        }
        
        // Populate tables
        populateBuildsTable();
        populateAddonsTable();
        
        console.log('Data loaded from localStorage successfully');
    } catch (e) {
        console.error('Error loading data from localStorage:', e);
        alert('There was an error loading your saved data. Falling back to default configuration.');
        
        // Fall back to config data if available
        if (typeof configData !== 'undefined') {
            loadDataFromConfig();
        }
    }
}

// Function to load data from the config file
function loadDataFromConfig() {
    // Store the configuration data
    buildData = configData.builds;
    installCodes = configData.installCodes;
    addonData = configData.addons;
    
    // Extract addon descriptions from config data if available
    if (configData.addonDescriptions) {
        addonDescriptionsData = configData.addonDescriptions;
    } else {
        // Otherwise, build addon descriptions from install codes
        addonDescriptionsData = {};
        installCodes.forEach(code => {
            if (code.forAddon) {
                addonDescriptionsData[code.forAddon] = code.description || '';
            }
        });
    }
    
    // Populate builds and add-ons tables
    populateBuildsTable();
    populateAddonsTable();
}

// Let's add a button to clear saved data if needed
function handleExcelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    uploadStatus.textContent = 'Reading file...';
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Process the Excel data
            processExcelData(workbook);
            
            uploadStatus.textContent = 'File processed successfully and data saved for future use!';
            uploadStatus.style.color = 'green';
            
            // Display notification that data is now saved
            alert('Your data has been successfully loaded and saved. You will not need to upload the file again unless you want to update the data.');
        } catch (error) {
            console.error('Error processing Excel file:', error);
            uploadStatus.textContent = 'Error processing file. Please check the format.';
            uploadStatus.style.color = 'red';
        }
    };
    
    reader.onerror = function() {
        uploadStatus.textContent = 'Error reading file.';
        uploadStatus.style.color = 'red';
    };
    
    reader.readAsArrayBuffer(file);
}

// Function to process Excel data and convert it to the required format
function processExcelData(workbook) {
    // Assume the first sheet contains the data
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert sheet to JSON
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (jsonData.length < 3) {
        throw new Error('Invalid Excel format. Not enough rows.');
    }
    
    // First row should have headers with "Code", "Description", "Type", and build numbers
    const headerRow = jsonData[0];
    
    // Find column indices
    const codeColIndex = headerRow.findIndex(cell => cell && cell.toString().trim().toLowerCase() === 'code');
    const descColIndex = headerRow.findIndex(cell => cell && cell.toString().trim().toLowerCase() === 'description');
    const typeColIndex = headerRow.findIndex(cell => cell && cell.toString().trim().toLowerCase() === 'type');
    
    if (codeColIndex === -1 || descColIndex === -1 || typeColIndex === -1) {
        throw new Error('Required columns (Code, Description, Type) not found in the Excel file.');
    }
    
    // Extract build numbers from header row (after the Type column)
    const buildNumbers = [];
    for (let i = typeColIndex + 1; i < headerRow.length; i++) {
        if (headerRow[i] && headerRow[i].toString().trim() !== '') {
            buildNumbers.push(headerRow[i].toString().trim());
        }
    }
    
    if (buildNumbers.length === 0) {
        throw new Error('No build numbers found in the Excel file.');
    }
    
    // Extract install codes and compatibility info
    const installCodesData = [];
    const addonsData = new Set();
    const addonDescriptionsObj = {};
    
    // Process data rows (starting from row 2)
    for (let rowIdx = 1; rowIdx < jsonData.length; rowIdx++) {
        const row = jsonData[rowIdx];
        
        // Skip empty rows
        if (!row || !row[codeColIndex] || row[codeColIndex].toString().trim() === '') {
            continue;
        }
        
        const code = row[codeColIndex].toString().trim();
        const description = row[descColIndex] ? row[descColIndex].toString().trim() : '';
        const type = row[typeColIndex] ? row[typeColIndex].toString().trim() : '';
        
        // Determine if this is an addon code
        const isAddon = type === 'Add-On';
        let addonName = null;
        
        if (isAddon) {
            // Use the code as the addon name for simplicity
            addonName = code;
            addonsData.add(addonName);
            // Store the description for the add-on
            addonDescriptionsObj[addonName] = description;
        }
        
        // Determine compatible builds
        const compatibleBuilds = [];
        for (let i = 0; i < buildNumbers.length; i++) {
            const cellValue = row[typeColIndex + 1 + i];
            // Any non-empty cell value (including checkboxes which might be TRUE/FALSE or 1/0)
            // Consider it compatible if the value exists and is not explicitly empty, "FALSE", "0", or "No"
            if (cellValue !== undefined && 
                cellValue !== null && 
                cellValue.toString().trim() !== '' && 
                cellValue.toString().trim().toLowerCase() !== 'false' && 
                cellValue.toString().trim() !== '0' && 
                cellValue.toString().trim().toLowerCase() !== 'no') {
                compatibleBuilds.push(buildNumbers[i]);
            }
        }
        
        // Create the install code object
        const installCodeObj = {
            code: code,
            description: description,
            compatibleBuilds: compatibleBuilds
        };
        
        if (isAddon) {
            installCodeObj.forAddon = addonName;
        }
        
        installCodesData.push(installCodeObj);
    }
    
    // Update the global data variables
    buildData = buildNumbers;
    installCodes = installCodesData;
    addonData = Array.from(addonsData);
    addonDescriptionsData = addonDescriptionsObj;
    
    // Save the data to localStorage for persistence
    saveDataToLocalStorage();
    
    // Reset any existing selections
    resetSelection();
    
    // Refresh the tables
    populateBuildsTable();
    populateAddonsTable();
    
    console.log('Processed Excel data:', {
        builds: buildData,
        installCodes: installCodes,
        addons: addonData,
        addonDescriptions: addonDescriptionsData
    });
}

// Function to save all data to localStorage
function saveDataToLocalStorage() {
    try {
        localStorage.setItem('buildData', JSON.stringify(buildData));
        localStorage.setItem('installCodes', JSON.stringify(installCodes));
        localStorage.setItem('addonData', JSON.stringify(addonData));
        localStorage.setItem('addonDescriptionsData', JSON.stringify(addonDescriptionsData));
        localStorage.setItem('lastUpdated', new Date().toString());
        console.log('Data saved to localStorage successfully');
    } catch (e) {
        console.error('Error saving data to localStorage:', e);
        alert('There was an error saving your data. This might be due to storage limits or privacy settings in your browser.');
    }
}

// Function to export the current data back to Excel
function exportToExcel() {
    // Create a new workbook
    const wb = XLSX.utils.book_new();
    
    // Prepare data for the worksheet
    const wsData = [];
    
    // Header row
    const headerRow = ['Code', 'Description', 'Type'];
    buildData.forEach(build => {
        headerRow.push(build);
    });
    wsData.push(headerRow);
    
    // Data rows for each install code
    installCodes.forEach(code => {
        const row = [];
        row.push(code.code);
        row.push(code.description);
        row.push(code.forAddon ? 'Add-On' : 'Required');
        
        // Mark compatible builds with an X
        buildData.forEach(build => {
            if (code.compatibleBuilds && code.compatibleBuilds.includes(build)) {
                row.push('X');
            } else {
                row.push('');
            }
        });
        
        wsData.push(row);
    });
    
    // Create worksheet
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'KTEC Flint Install Codes');
    
    // Generate Excel file
    XLSX.writeFile(wb, 'KTEC_Flint_Install_Codes.xlsx');
}

// Function to populate builds table with filtering
function populateBuildsTable() {
    const buildsTableBody = document.querySelector('#builds-table tbody');
    buildsTableBody.innerHTML = '';
    
    // Get filter value
    const searchTerm = buildSearchInput.value.toLowerCase();
    
    // Filter and populate builds
    const filteredBuilds = buildData.filter(build => 
        build.toLowerCase().includes(searchTerm)
    );
    
    if (filteredBuilds.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = '<td colspan="2">No builds found matching your search.</td>';
        buildsTableBody.appendChild(row);
        return;
    }
    
    filteredBuilds.forEach(build => {
        const row = document.createElement('tr');
        
        row.innerHTML = `
            <td>${build}</td>
            <td>
                <button class="select-build" data-build="${build}">
                    Select
                </button>
            </td>
        `;
        
        buildsTableBody.appendChild(row);
    });
    
    // Add event listeners to select buttons
    document.querySelectorAll('.select-build').forEach(button => {
        button.addEventListener('click', function() {
            const buildNumber = this.getAttribute('data-build');
            selectBuild(buildNumber);
        });
    });
}

// Function to populate add-ons table with filtering
function populateAddonsTable() {
    const addonsTableBody = document.querySelector('#addons-table tbody');
    addonsTableBody.innerHTML = '';
    
    // Get filter value
    const searchTerm = addonSearchInput.value.toLowerCase();
    
    // Filter and populate add-ons
    const filteredAddons = addonData.filter(addon => 
        addon.toLowerCase().includes(searchTerm) || 
        (addonDescriptionsData[addon] && addonDescriptionsData[addon].toLowerCase().includes(searchTerm))
    );
    
    if (filteredAddons.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = '<td colspan="3">No add-ons found matching your search.</td>';
        addonsTableBody.appendChild(row);
        return;
    }
    
    filteredAddons.forEach(addon => {
        const row = document.createElement('tr');
        row.className = 'addon-row';
        
        // Check if this add-on is already selected
        const isSelected = selectedAddons.includes(addon);
        if (isSelected) {
            row.classList.add('selected');
        }
        
        // Get description for this add-on
        const description = addonDescriptionsData[addon] || '';
        
        row.innerHTML = `
            <td>${addon}</td>
            <td>${description}</td>
            <td>
                <button class="${isSelected ? 'reset-button' : 'select-build'}" data-addon="${addon}">
                    ${isSelected ? 'Remove' : 'Add'}
                </button>
            </td>
        `;
        
        addonsTableBody.appendChild(row);
    });
    
    // Add event listeners to add/remove buttons
    document.querySelectorAll('#addons-table button').forEach(button => {
        button.addEventListener('click', function() {
            const addon = this.getAttribute('data-addon');
            if (selectedAddons.includes(addon)) {
                removeAddon(addon);
            } else {
                addAddon(addon);
            }
        });
    });
}

// Function to add an add-on
function addAddon(addon) {
    // Add to selected addons array if not already there
    if (!selectedAddons.includes(addon)) {
        selectedAddons.push(addon);
    }
    
    // Update the selected add-ons list
    updateSelectedAddonsList();
    
    // Update the add-ons table to reflect the change
    populateAddonsTable();
}

// Function to remove an add-on
function removeAddon(addon) {
    // Remove from selected addons array
    selectedAddons = selectedAddons.filter(item => item !== addon);
    
    // Update the selected add-ons list
    updateSelectedAddonsList();
    
    // Update the add-ons table to reflect the change
    populateAddonsTable();
}

// Function to update the selected add-ons list display
function updateSelectedAddonsList() {
    if (selectedAddons.length === 0) {
        selectedAddonsList.innerHTML = '<p class="no-addons">No add-ons selected</p>';
        return;
    }
    
    selectedAddonsList.innerHTML = '';
    
    selectedAddons.forEach(addon => {
        const addonItem = document.createElement('div');
        addonItem.className = 'addon-item';
        
        // Get the addon notes or create new ones from the description
        // If there are no saved notes yet, use the addon description as a starting point
        if (!addonDescriptions[addon] && addonDescriptionsData[addon]) {
            addonDescriptions[addon] = addonDescriptionsData[addon];
            
            // Save the initial description to localStorage
            try {
                localStorage.setItem('addonCustomDescriptions', JSON.stringify(addonDescriptions));
            } catch (e) {
                console.error('Error saving addon descriptions:', e);
            }
        }
        
        let addonNotes = addonDescriptions[addon] || '';
        
        addonItem.innerHTML = `
            <h4>${addon}</h4>
            <button class="remove-addon" data-addon="${addon}">×</button>
            <div class="editable-description">
                <label for="addon-notes-${addon.replace(/\s+/g, '-')}"><strong>Notes:</strong></label>
                <textarea id="addon-notes-${addon.replace(/\s+/g, '-')}" 
                    data-addon="${addon}" 
                    placeholder="Add your notes for this add-on...">${addonNotes}</textarea>
            </div>
        `;
        
        selectedAddonsList.appendChild(addonItem);
    });
    
    // Add event listeners to remove buttons
    document.querySelectorAll('.remove-addon').forEach(button => {
        button.addEventListener('click', function() {
            const addon = this.getAttribute('data-addon');
            removeAddon(addon);
        });
    });
    
    // Add event listeners to textarea changes
    document.querySelectorAll('[id^="addon-notes-"]').forEach(textarea => {
        textarea.addEventListener('input', function() {
            const addon = this.getAttribute('data-addon');
            addonDescriptions[addon] = this.value;
            
            // Save to localStorage
            try {
                localStorage.setItem('addonCustomDescriptions', JSON.stringify(addonDescriptions));
            } catch (e) {
                console.error('Error saving addon descriptions:', e);
            }
        });
    });
}

// Function to select a build and show compatible install codes
function selectBuild(buildNumber) {
    // Set selected build info
    selectedBuildNumber.textContent = buildNumber;
    
    // Check for custom description
    if (customDescriptions[buildNumber]) {
        buildDescriptionTextarea.value = customDescriptions[buildNumber];
    } else {
        buildDescriptionTextarea.value = '';
    }
    
    // Populate install codes table
    populateInstallCodesTable(buildNumber);
    
    // Show result container
    resultContainer.style.display = 'block';
    
    // Highlight selected build in the table
    document.querySelectorAll('#builds-table tbody tr').forEach(row => {
        row.classList.remove('highlight');
        
        const rowBuildNumber = row.cells[0].textContent;
        if (rowBuildNumber === buildNumber) {
            row.classList.add('highlight');
        }
    });
    
    // Scroll to results
    resultContainer.scrollIntoView({ behavior: 'smooth' });
}

// Function to populate install codes table based on selected build
function populateInstallCodesTable(buildNumber) {
    const codesTableBody = document.querySelector('#codes-table tbody');
    codesTableBody.innerHTML = '';
    
    // Filter to only show required codes that are compatible with the selected build
    const requiredCodes = installCodes.filter(code => 
        !code.forAddon && // Not an addon code
        code.compatibleBuilds && 
        code.compatibleBuilds.includes(buildNumber)
    );
    
    if (requiredCodes.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = '<td colspan="3">No required install codes found for this build.</td>';
        codesTableBody.appendChild(row);
        return;
    }
    
    requiredCodes.forEach(code => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${code.code}</td>
            <td>${code.description}</td>
            <td class="compatible">✓ Required</td>
        `;
        
        // Add highlight class to all rows since they're all required
        row.classList.add('highlight');
        
        codesTableBody.appendChild(row);
    });
    
    // Apply any existing filter
    filterInstallCodes();
}

// Function to filter install codes based on search input
function filterInstallCodes() {
    const searchTerm = codeSearchInput.value.toLowerCase();
    const rows = document.querySelectorAll('#codes-table tbody tr');
    
    rows.forEach(row => {
        const code = row.cells[0].textContent.toLowerCase();
        const description = row.cells[1].textContent.toLowerCase();
        const isMatch = code.includes(searchTerm) || description.includes(searchTerm);
        
        row.style.display = isMatch ? '' : 'none';
    });
}

// Function to reset selection
function resetSelection() {
    // Hide result container
    resultContainer.style.display = 'none';
    
    // Clear search inputs
    buildSearchInput.value = '';
    codeSearchInput.value = '';
    addonSearchInput.value = '';
    
    // Clear selected add-ons
    selectedAddons = [];
    updateSelectedAddonsList();
    
    // Remove highlights
    document.querySelectorAll('#builds-table tbody tr').forEach(row => {
        row.classList.remove('highlight');
    });
    
    // Refresh builds table
    populateBuildsTable();
    populateAddonsTable();
}

// Function to prepare the print view
function preparePrintView() {
    // Check if a build is selected
    const currentBuild = selectedBuildNumber.textContent;
    if (!currentBuild) {
        alert('Please select a build first.');
        return;
    }
    
    // Format the date
    const now = new Date();
    const formattedDate = now.toLocaleDateString('en-US', {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
    const formattedTime = now.toLocaleTimeString('en-US', {
        hour: '2-digit',
        minute: '2-digit'
    });
    
    // Fill in the print template
    printBuildId.textContent = `Build: ${currentBuild}`;
    printDate.textContent = `Date: ${formattedDate}`;
    
    // Get notes
    const notes = buildDescriptionTextarea.value.trim();
    if (notes) {
        printNotes.innerHTML = `<p><strong>Notes:</strong></p><p>${notes}</p>`;
    } else {
        printNotes.innerHTML = '<p><em>No notes added.</em></p>';
    }
    
    // Clear and fill in the compatible codes table
    printCodesTable.innerHTML = '';
    const compatibleCodes = [];
    
    // Find compatible install codes
    installCodes.forEach(code => {
        if (code.forAddon) return; // Skip addon codes
        if (code.compatibleBuilds && code.compatibleBuilds.includes(currentBuild)) {
            compatibleCodes.push(code);
        }
    });
    
    if (compatibleCodes.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = '<td colspan="2">No compatible install codes found.</td>';
        printCodesTable.appendChild(row);
    } else {
        compatibleCodes.forEach(code => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${code.code}</td>
                <td>${code.description}</td>
            `;
            printCodesTable.appendChild(row);
        });
    }
    
    // Process add-ons
    printAddonsList.innerHTML = '';
    
    if (selectedAddons.length === 0) {
        printAddonsSection.style.display = 'none';
    } else {
        printAddonsSection.style.display = 'block';
        
        selectedAddons.forEach(addon => {
            const addonItem = document.createElement('div');
            addonItem.className = 'print-addon-item';
            
            // Find install code for this add-on
            const addonCode = installCodes.find(code => code.forAddon === addon);
            const codeText = addonCode ? addonCode.code : 'N/A';
            const codeDesc = addonCode ? addonCode.description : 'No description available';
            
            // Get notes for this add-on
            const addonNotes = addonDescriptions[addon] || '';
            
            // Build the add-on item HTML
            addonItem.innerHTML = `
                <div class="print-addon-title">${addon}</div>
                <div class="print-addon-codes">
                    <table>
                        <tr>
                            <th>Install Code</th>
                            <th>Description</th>
                        </tr>
                        <tr>
                            <td>${codeText}</td>
                            <td>${codeDesc}</td>
                        </tr>
                    </table>
                </div>
                ${addonNotes ? `
                    <div class="print-addon-notes">
                        <p><strong>Notes:</strong></p>
                        <p>${addonNotes}</p>
                    </div>
                ` : ''}
            `;
            
            printAddonsList.appendChild(addonItem);
        });
    }
    
    // Set timestamp
    printTimestamp.textContent = `${formattedDate} at ${formattedTime}`;
    
    // Show the print container
    printContainer.style.display = 'block';
}