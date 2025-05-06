// KTEC Flint Install Code Generator - Default Configuration
// This file provides starter data until you upload your Excel file

const configData = {
    // Basic list of build numbers to start with
    builds: [
        "KF-DEMO-1",
        "KF-DEMO-2",
        "KF-DEMO-3"
    ],
    
    // Sample install codes
    installCodes: [
        {
            code: "IC-DEMO-1",
            description: "Demo Installation Package",
            compatibleBuilds: ["KF-DEMO-1", "KF-DEMO-2"]
        },
        {
            code: "IC-DEMO-2",
            description: "Demo Configuration",
            compatibleBuilds: ["KF-DEMO-1", "KF-DEMO-2", "KF-DEMO-3"]
        },
        {
            code: "IC-DEMO-A1",
            description: "Demo Add-On Installation",
            forAddon: "Demo Add-On",
            compatibleBuilds: ["KF-DEMO-1", "KF-DEMO-2"]
        }
    ],
    
    // Sample add-ons
    addons: [
        "Demo Add-On"
    ],
    
    // Add-on descriptions (optional, will be extracted from install codes if not provided)
    addonDescriptions: {
        "Demo Add-On": "This is a demonstration add-on with sample functionality"
    }
};