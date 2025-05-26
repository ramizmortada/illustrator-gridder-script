#target illustrator

//
// 1. STUB UNDO METHODS FOR OLDER ILLUSTRATOR (CS6/CC2014...)
// If Illustrator lacks these, they become no-ops.
//
if (typeof app.beginUndoGroup !== "function") {
    app.beginUndoGroup = function(groupName) { /* no-op */ };
    app.endUndoGroup   = function() { /* no-op */ };
}

(function main() {
    // Must have at least one open document
    if (app.documents.length === 0) {
        alert("Please open a document before running the script.");
        return;
    }
    
    // Constants and variables
    var doc = app.activeDocument;
    var isUndo = false; // Whether we have a preview to undo
    var DEBUG = true; // Set to true to display debug information
    
    // Settings file path
    var SETTINGS_FOLDER = (function() {
        var folder;
        if (Folder.fs === "Windows") {
            folder = new Folder(Folder.userData + "/GridderScript");
        } else {
            folder = new Folder(Folder.userData + "/GridderScript");
        }
        
        if (!folder.exists) {
            folder.create();
        }
        return folder;
    })();
    
    var SETTINGS_FILE = File(SETTINGS_FOLDER + "/GridderSettings.jsx");

    function debugLog(message) {
        if (DEBUG) {
            $.writeln(message);
        }
    }
    
    // Default settings
    var DEFAULT_SETTINGS = {
        rows: "4",
        cols: "4",
        margin: "20",
        marginUnit: 0, // index for "pt"
        gutter: "10",
        gutterUnit: 0, // index for "pt"
        applyAll: true
    };
    
    // Load settings
    function loadSettings() {
        var settings = DEFAULT_SETTINGS; // Start with defaults
        
        try {
            if (SETTINGS_FILE.exists) {
                debugLog("Settings file exists at: " + SETTINGS_FILE.fsName);
                
                // Try to load and evaluate the settings file
                SETTINGS_FILE.encoding = "UTF8";
                if (SETTINGS_FILE.open("r")) {
                    var settingsCode = SETTINGS_FILE.read();
                    SETTINGS_FILE.close();
                    
                    if (settingsCode && settingsCode.length > 0) {
                        // Create a function to evaluate the settings safely
                        var getSettingsFunction = new Function("return " + settingsCode + ";");
                        var loadedSettings = getSettingsFunction();
                        
                        if (loadedSettings && typeof loadedSettings === "object") {
                            debugLog("Settings loaded successfully");
                            
                            // Merge with defaults to ensure all fields exist
                            for (var key in loadedSettings) {
                                if (DEFAULT_SETTINGS.hasOwnProperty(key)) {
                                    settings[key] = loadedSettings[key];
                                }
                            }
                        } else {
                            debugLog("Invalid settings format, using defaults");
                        }
                    } else {
                        debugLog("Settings file is empty, using defaults");
                    }
                } else {
                    debugLog("Could not open settings file, using defaults");
                }
            } else {
                debugLog("Settings file does not exist, using defaults");
                // Create with defaults on first run
                saveSettings(DEFAULT_SETTINGS);
            }
        } catch (e) {
            debugLog("Error loading settings: " + e);
            alert("Settings could not be loaded. Using defaults.\nError: " + e);
        }
        
        return settings;
    }
    
    // Save settings
    function saveSettings(settings) {
        try {
            debugLog("Saving settings to: " + SETTINGS_FILE.fsName);
            
            // Format settings as JavaScript object literal
            var settingsCode = "{\n";
            for (var key in settings) {
                if (settings.hasOwnProperty(key)) {
                    var value = settings[key];
                    
                    // Format based on value type
                    if (typeof value === "string") {
                        settingsCode += "    " + key + ": \"" + value + "\",\n";
                    } else if (typeof value === "number") {
                        settingsCode += "    " + key + ": " + value + ",\n";
                    } else if (typeof value === "boolean") {
                        settingsCode += "    " + key + ": " + value + ",\n";
                    }
                }
            }
            settingsCode += "}";
            
            // Write to the settings file
            SETTINGS_FILE.encoding = "UTF8";
            if (SETTINGS_FILE.open("w")) {
                SETTINGS_FILE.write(settingsCode);
                SETTINGS_FILE.close();
                debugLog("Settings saved successfully");
                return true;
            } else {
                debugLog("Failed to open settings file for writing");
                return false;
            }
        } catch (e) {
            debugLog("Error saving settings: " + e);
            alert("Settings could not be saved.\nError: " + e);
            return false;
        }
    }

    // -----------------------------------------------
    // CREATE DIALOG
    // -----------------------------------------------
    var dialog = new Window("dialog", "Gridder");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    // Load existing settings
    var currentSettings = loadSettings();
    
    if (DEBUG) {
        $.writeln("Loaded settings: " + $.global["GridderSettingsStr"]);
    }

    // BRANDING
    var brandingGroup = dialog.add("group");
    brandingGroup.alignment = "center";
    brandingGroup.add("statictext", undefined, "Made by: Ahmed Ali Yousef");
    
    // ROWS
    dialog.add("statictext", undefined, "Rows:");
    var rowsInput = dialog.add("edittext", undefined, currentSettings.rows);
    rowsInput.characters = 5;
    setupArrowsForInteger(rowsInput, 1);
    setupMouseWheelForInteger(rowsInput, 1);

    // COLUMNS
    dialog.add("statictext", undefined, "Columns:");
    var colsInput = dialog.add("edittext", undefined, currentSettings.cols);
    colsInput.characters = 5;
    setupArrowsForInteger(colsInput, 1);
    setupMouseWheelForInteger(colsInput, 1);

    // MARGIN
    dialog.add("statictext", undefined, "Margin:");
    var marginGroup = dialog.add("group");
    marginGroup.orientation = "row";
    marginGroup.alignChildren = "fill";

    var marginInput = marginGroup.add("edittext", undefined, currentSettings.margin);
    marginInput.characters = 5;
    setupArrowsForFloat(marginInput, 0);
    setupMouseWheelForFloat(marginInput, 0);

    var marginUnitDropdown = marginGroup.add("dropdownlist", undefined, ["pt","px","mm","cm","in"]);
    try {
        marginUnitDropdown.selection = parseInt(currentSettings.marginUnit) || 0;
    } catch (e) {
        marginUnitDropdown.selection = 0; // Default to "pt"
        debugLog("Error setting margin unit: " + e);
    }

    // GUTTER
    dialog.add("statictext", undefined, "Gutter:");
    var gutterGroup = dialog.add("group");
    gutterGroup.orientation = "row";
    gutterGroup.alignChildren = "fill";

    var gutterInput = gutterGroup.add("edittext", undefined, currentSettings.gutter);
    gutterInput.characters = 5;
    setupArrowsForFloat(gutterInput, 0);
    setupMouseWheelForFloat(gutterInput, 0);

    var gutterUnitDropdown = gutterGroup.add("dropdownlist", undefined, ["pt","px","mm","cm","in"]);
    try {
        gutterUnitDropdown.selection = parseInt(currentSettings.gutterUnit) || 0;
    } catch (e) {
        gutterUnitDropdown.selection = 0; // Default to "pt" 
        debugLog("Error setting gutter unit: " + e);
    }

    // APPLY TO ALL ARTBOARDS
    var applyAllCheck = dialog.add("checkbox", undefined, "Apply to all artboards");
    applyAllCheck.value = !!currentSettings.applyAll; // Convert to boolean

    // PREVIEW, OK, CANCEL
    var btnGroup = dialog.add("group");
    btnGroup.alignment = "center";

    var previewCheck = btnGroup.add("checkbox", undefined, "Preview");
    previewCheck.value = true; // default ON

    var cancelBtn = btnGroup.add("button", undefined, "Cancel");
    var okBtn = btnGroup.add("button", undefined, "OK");
    
    // Add help button
    var helpBtn = btnGroup.add("button", undefined, "Help");
    helpBtn.onClick = function() {
        var helpText = "Gridder Settings:\n\n";
        helpText += "Settings are saved to:\n" + SETTINGS_FILE.fsName + "\n\n";
        helpText += "Current values:\n";
        helpText += "- Rows: " + rowsInput.text + "\n";
        helpText += "- Cols: " + colsInput.text + "\n";
        helpText += "- Margin: " + marginInput.text + " " + marginUnitDropdown.selection.text + "\n";
        helpText += "- Gutter: " + gutterInput.text + " " + gutterUnitDropdown.selection.text + "\n";
        
        alert(helpText);
    };

    // -----------------------------------------------
    // EVENT HANDLERS
    // -----------------------------------------------
    rowsInput.onChanging      = function() { if (previewCheck.value) previewStart(); };
    colsInput.onChanging      = function() { if (previewCheck.value) previewStart(); };
    marginInput.onChanging    = function() { if (previewCheck.value) previewStart(); };
    gutterInput.onChanging    = function() { if (previewCheck.value) previewStart(); };

    marginUnitDropdown.onChange = function() { if (previewCheck.value) previewStart(); };
    gutterUnitDropdown.onChange = function() { if (previewCheck.value) previewStart(); };

    applyAllCheck.onClick = function() { if (previewCheck.value) previewStart(); };
    previewCheck.onClick  = function() { previewStart(); };

    cancelBtn.onClick = function() { handleCancel(); };
    okBtn.onClick     = function() { handleOk(); };

    // Show the dialog
    dialog.center();
    dialog.show();

    // --------------------------------------------------
    // PREVIEW LOGIC
    // --------------------------------------------------
    function previewStart() {
        try {
            if (previewCheck.value) {
                // If there's a previous preview, undo it
                if (isUndo) {
                    app.undo();
                    isUndo = false;
                }
                // Create new preview
                app.beginUndoGroup("Grid Preview");
                var didDraw = createGridAsGuides();
                app.endUndoGroup();
                app.redraw();

                if (didDraw) {
                    isUndo = true;
                }
            } else {
                // If PREVIEW is OFF, remove existing preview if any
                if (isUndo) {
                    app.undo();
                    app.redraw();
                    isUndo = false;
                }
            }
        } catch (err) {
            alert("Preview error:\n" + err + "\nLine: " + (err.line || "unknown"));
        }
    }

    function handleOk() {
        if (previewCheck.value && isUndo) {
            // If there's a valid preview, keep it
            isUndo = false;
        } else {
            // Otherwise, create guides once
            app.beginUndoGroup("Final Grid");
            createGridAsGuides();
            app.endUndoGroup();
        }

        // Save settings
        var settingsToSave = {
            rows: rowsInput.text,
            cols: colsInput.text,
            margin: marginInput.text,
            marginUnit: marginUnitDropdown.selection.index,
            gutter: gutterInput.text,
            gutterUnit: gutterUnitDropdown.selection.index,
            applyAll: applyAllCheck.value
        };
        
        var saveResult = saveSettings(settingsToSave);
        // Silent save - no alert needed
        
        dialog.close();
    }

    function handleCancel() {
        // If there's a preview, remove it
        if (isUndo) {
            app.undo();
            app.redraw();
            isUndo = false;
        }
        dialog.close();
    }

    // --------------------------------------------------
    // MAIN CREATE FUNCTION
    // --------------------------------------------------
    function createGridAsGuides() {
        var rowsVal   = parseInt(rowsInput.text, 10);
        var colsVal   = parseInt(colsInput.text, 10);

        var marginVal = parseFloat(marginInput.text);
        var gutterVal = parseFloat(gutterInput.text);

        var marginPts = convertToPoints(marginVal, marginUnitDropdown.selection.text);
        var gutterPts = convertToPoints(gutterVal, gutterUnitDropdown.selection.text);

        if (!validateNumbers(rowsVal, colsVal, marginPts, gutterPts)) {
            return false;
        }

        if (applyAllCheck.value) {
            for (var i = 0; i < doc.artboards.length; i++) {
                drawGuidesOnArtboard(doc.artboards[i], rowsVal, colsVal, marginPts, gutterPts);
            }
        } else {
            var activeAB = doc.artboards[doc.artboards.getActiveArtboardIndex()];
            drawGuidesOnArtboard(activeAB, rowsVal, colsVal, marginPts, gutterPts);
        }
        return true;
    }

    function validateNumbers(r, c, margin, gutter) {
        if (isNaN(r) || r < 1 || isNaN(c) || c < 1) {
            alert("Rows and Columns must be 1 or higher.");
            return false;
        }
        if (isNaN(margin) || margin < 0 || isNaN(gutter) || gutter < 0) {
            alert("Margin/Gutter must be non-negative numbers.");
            return false;
        }
        return true;
    }

    // --------------------------------------------------
    // DRAW GUIDES ON A SINGLE ARTBOARD
    // --------------------------------------------------
    function drawGuidesOnArtboard(artboard, rows, cols, margin, gutter) {
        var bounds  = artboard.artboardRect; // [left, top, right, bottom]
        var abWidth  = bounds[2] - bounds[0];
        var abHeight = bounds[1] - bounds[3];

        var layer = getOrCreateLayer("Gridder_Guides");

        var usableWidth  = abWidth  - 2 * margin;
        var usableHeight = abHeight - 2 * margin;
        var totalGutterW = (cols - 1) * gutter;
        var totalGutterH = (rows - 1) * gutter;

        var colWidth  = (usableWidth  - totalGutterW) / cols;
        var rowHeight = (usableHeight - totalGutterH) / rows;

        // 1) Margin rectangle
        var marginRect = layer.pathItems.rectangle(
            bounds[1] - margin,   // top
            bounds[0] + margin,   // left
            usableWidth,
            usableHeight
        );
        makeGuide(marginRect);

        // 2) Vertical gutters
        for (var cIndex = 1; cIndex < cols; cIndex++) {
            var xPos = bounds[0] + margin + cIndex * colWidth + (cIndex - 1) * gutter;
            var gutterRect = layer.pathItems.rectangle(
                bounds[1] - margin,
                xPos,
                gutter,
                usableHeight
            );
            makeGuide(gutterRect);
        }

        // 3) Horizontal gutters
        for (var rIndex = 1; rIndex < rows; rIndex++) {
            var yPos = bounds[1] - margin - rIndex * rowHeight - (rIndex - 1) * gutter;
            var gutterRectH = layer.pathItems.rectangle(
                yPos,
                bounds[0] + margin,
                usableWidth,
                gutter
            );
            makeGuide(gutterRectH);
        }
    }

    function convertToPoints(value, unit) {
        if (isNaN(value)) return NaN;

        switch(unit) {
            case "px":
            case "pt": 
                return value;
            case "mm":
                return value * 2.83464566929134; 
            case "cm":
                return value * 28.3464566929134;
            case "in":
                return value * 72;
            default:
                return value;
        }
    }

    function makeGuide(item) {
        item.stroked = false;
        item.filled  = false;
        item.guides  = true;
    }

    function getOrCreateLayer(layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                return doc.layers[i];
            }
        }
        var newLayer = doc.layers.add();
        newLayer.name = layerName;
        return newLayer;
    }

    // --------------------------------------------------
    // UTILITY: SETUP ARROW KEYS FOR INTEGER FIELDS
    // (Up/Down => +/-1, Shift+Up/Down => +/-10)
    // --------------------------------------------------
    function setupArrowsForInteger(editField, minValue) {
        editField.addEventListener("keydown", function(e) {
            if (e.keyName === "Up" || e.keyName === "Down") {
                var step = e.shiftKey ? 10 : 1;
                var val  = parseInt(editField.text, 10);
                if (isNaN(val)) val = minValue; // Default to minValue if NaN

                if (e.keyName === "Up") {
                    val += step;
                } else {
                    val -= step;
                }
                if (val < minValue) val = minValue; // clamp at minValue

                editField.text = val.toString();

                if (typeof e.preventDefault === "function") e.preventDefault();
                else e.returnValue = false;

                if (previewCheck.value) previewStart();
            }
        });
    }

    // --------------------------------------------------
    // UTILITY: SETUP ARROW KEYS FOR FLOAT FIELDS
    // (Up/Down => +/-1, Shift => +/-10, clamp to >=0)
    // --------------------------------------------------
    function setupArrowsForFloat(editField, minValue) {
        editField.addEventListener("keydown", function(e) {
            if (e.keyName === "Up" || e.keyName === "Down") {
                var step = e.shiftKey ? 10 : 1;
                var val  = parseFloat(editField.text);
                if (isNaN(val)) val = minValue; // Default to minValue if NaN

                if (e.keyName === "Up") {
                    val += step;
                } else {
                    val -= step;
                }
                if (val < minValue) val = minValue; // clamp at minValue

                editField.text = val.toString();

                if (typeof e.preventDefault === "function") e.preventDefault();
                else e.returnValue = false;

                if (previewCheck.value) previewStart();
            }
        });
    }

    // --------------------------------------------------
    // UTILITY: SETUP MOUSE WHEEL FOR INTEGER FIELDS
    // (Wheel => +/-1, Shift => +/-10)
    // --------------------------------------------------
    function setupMouseWheelForInteger(editField, minValue) {
        editField.addEventListener("mousewheel", function(e) {
            // Some environments won't fire this event at all
            // or it might be "wheelDelta" vs. "detail" depending on OS
            var val = parseInt(editField.text, 10);
            if (isNaN(val)) val = minValue; // Default to minValue if NaN

            var step = e.shiftKey ? 10 : 1;

            // "wheelDelta" is positive on wheel up, negative on wheel down in many Windows setups.
            // Some macOS setups invert it. There's no universal standard in ScriptUI.
            if (e.wheelDelta > 0) {
                val += step;
            } else {
                val -= step;
            }
            if (val < minValue) val = minValue; // clamp at minValue

            editField.text = val.toString();

            if (typeof e.preventDefault === "function") e.preventDefault();
            else e.returnValue = false;

            if (previewCheck.value) previewStart();
        });
    }

    // --------------------------------------------------
    // UTILITY: SETUP MOUSE WHEEL FOR FLOAT FIELDS
    // (Wheel => +/-1, Shift => +/-10)
    // --------------------------------------------------
    function setupMouseWheelForFloat(editField, minValue) {
        editField.addEventListener("mousewheel", function(e) {
            var val = parseFloat(editField.text);
            if (isNaN(val)) val = minValue; // Default to minValue if NaN

            var step = e.shiftKey ? 10 : 1;

            if (e.wheelDelta > 0) {
                val += step;
            } else {
                val -= step;
            }
            if (val < minValue) val = minValue; // clamp at minValue

            editField.text = val.toString();

            if (typeof e.preventDefault === "function") e.preventDefault();
            else e.returnValue = false;

            if (previewCheck.value) previewStart();
        });
    }

})();
