// NAME: Top Expand
// STATUS: Ready
// FUNCTION: Increase height of selected items from top one baseline for each time the script is run
// Best to use on non group grouped items for Nav. If needed use white arrow to handle groups, but black arrow is find for text boxes
// Prefered keyboard shortcut is Top Expand F9
// Version 1.0.0



//Top Expand
// Define the script as a function to be executed
app.doScript(function() {
    // Get the active document
    var doc = app.activeDocument;

    // Get document grid preferences
    var gridPreferences = doc.gridPreferences;
    var baselineIncrement = gridPreferences.baselineDivision;

    // Function to adjust height from the top upwards
    function adjustHeight(item) {
        var currentHeight = item.geometricBounds[2] - item.geometricBounds[0];
        
        if (currentHeight <= baselineIncrement) {
            // If the height is less than or equal to the baseline increment, move the item upwards
            var direction = (baselineIncrement > 0) ? -1 : 1; // Text flows down or up
            
            // Adjust geometric bounds to move upwards by one baseline increment
            item.geometricBounds = [
                item.geometricBounds[0] + direction * baselineIncrement,  // Top edge Y-coordinate moves upwards
                item.geometricBounds[1],  // Left edge X-coordinate remains the same
                item.geometricBounds[2] + direction * baselineIncrement, // Bottom edge moves upwards
                item.geometricBounds[3]  // Right edge X-coordinate remains the same
            ];
        } else {
            // If the height is greater than the baseline increment, adjust normally
            item.geometricBounds = [
                item.geometricBounds[0] - baselineIncrement, // Adjust top edge upwards
                item.geometricBounds[1], 
                item.geometricBounds[2], 
                item.geometricBounds[3]
            ];
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
        } else if (item instanceof Image) {
            // Check if the image is inside a rectangle (frame)
            var parent = item.parent;
            if (parent instanceof Rectangle) {
                adjustHeight(parent); // Adjust the frame's height, not the image itself
            } else {
                // For standalone images not inside a rectangle, adjust the image's bounds directly
                adjustHeight(item);
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
