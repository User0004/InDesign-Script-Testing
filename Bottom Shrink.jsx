// NAME: Bottom Shrink
// STATUS: Ready
// FUNCTION: Decrease height of selected items from top one baseline for each time the script is run
// Best to use on non group grouped items for Nav. If needed use white arrow to handle groups, but black arrow is find for text boxes 
// Prefered keyboard shortcut is Bottom Shrink shift F10

// Bottom shrink
// USE from here 

app.doScript(function() {
    // Get the active document
    var doc = app.activeDocument;

    // Get document grid preferences
    var gridPreferences = doc.gridPreferences;
    var baselineIncrement = gridPreferences.baselineDivision;

    // Add a precision threshold to handle floating-point precision issues
    var precisionThreshold = 0.1; // This helps to avoid errors when height is close to baselineIncrement

    // Function to adjust height
    function adjustHeight(item) {
        var currentHeight = item.geometricBounds[2] - item.geometricBounds[0];
        
        if (currentHeight <= baselineIncrement + precisionThreshold) {
            // If the height is less than or equal to the baseline increment (with tolerance), move the item in the direction of text flow
            var direction = (baselineIncrement > 0) ? 1 : -1; // Determine if text flows down or up
            var moveAmount = Math.abs(baselineIncrement); // Absolute value of baseline increment
            
            // Adjust geometric bounds based on text flow direction
            item.geometricBounds = [
                item.geometricBounds[0] - direction * moveAmount,
                item.geometricBounds[1], // Left edge X-coordinate remains the same
                item.geometricBounds[2] - direction * moveAmount,
                item.geometricBounds[3]  // Right edge X-coordinate remains the same
            ];
        } else if (currentHeight > baselineIncrement + precisionThreshold) {
            // If the height is greater than the baseline increment, reduce height from the bottom
            if (item instanceof Image && item.parent instanceof Rectangle) {
                // Adjust the height of the parent frame, not the image itself
                adjustHeight(item.parent);
            } else {
                var newHeight = currentHeight - baselineIncrement;
                item.geometricBounds = [
                    item.geometricBounds[0], 
                    item.geometricBounds[1], 
                    item.geometricBounds[2] - baselineIncrement, 
                    item.geometricBounds[3]
                ];
            }
        }
    }

    // Function to process each item or group recursively
    function processItem(item) {
        if (item instanceof GraphicLine || item instanceof TextFrame || item instanceof Rectangle) {
            adjustHeight(item);
        } else if (item instanceof Group) {
            for (var j = 0; j < item.allPageItems.length; j++) {
                processItem(item.allPageItems[j]);
            }
        } else {
            alert("Selection contains unsupported items. Please select only vertical rules, text frames, picture boxes, or groups.");
        }
    }

    // Check if there are any selected objects
    if (app.selection.length === 0) {
        alert("Please select one or more vertical rules, text frames, picture boxes, or groups.");
    } else {
        // Iterate over all selected objects
        for (var i = 0; i < app.selection.length; i++) {
            processItem(app.selection[i]);
        }
    }
}, ScriptLanguage.JAVASCRIPT, null, UndoModes.ENTIRE_SCRIPT, "Adjust Item Height Based on Baseline Increment");
