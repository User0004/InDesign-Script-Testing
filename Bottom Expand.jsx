// NAME: Bottom Expand
// STATUS: Ready
// FUNCTION: Increase height of selected items from bottom one baseline for each time the script is run
// Best to use on non group grouped items for Nav. If needed use white arrow to handle groups, but black arrow is find for text boxes 
// Prefered keyboard shortcut is Bottom Expand F10


//Bottom Expand
app.doScript(function() {
    // Get the active document
    var doc = app.activeDocument;

    // Get document grid preferences
    var gridPreferences = doc.gridPreferences;
    var baselineIncrement = gridPreferences.baselineDivision;

    // Function to adjust height
    function adjustHeight(item) {
        var currentHeight = item.geometricBounds[2] - item.geometricBounds[0];
        
        if (currentHeight <= baselineIncrement) {
            // If the height is less than or equal to the baseline increment, move the item in the direction of text flow
            var direction = (baselineIncrement > 0) ? 1 : -1; // Determine if text flows down or up
            var moveAmount = Math.abs(baselineIncrement); // Absolute value of baseline increment
            
            // Adjust geometric bounds based on text flow direction
            item.geometricBounds = [
                item.geometricBounds[0] + direction * moveAmount,
                item.geometricBounds[1], // Left edge X-coordinate remains the same
                item.geometricBounds[2] + direction * moveAmount, // Expand height from the bottom
                item.geometricBounds[3]  // Right edge X-coordinate remains the same
            ];
        } else {
            // If the height is greater than the baseline increment, expand height from the bottom
            var newHeight = currentHeight + baselineIncrement;
            item.geometricBounds = [
                item.geometricBounds[0], 
                item.geometricBounds[1], 
                item.geometricBounds[2] + baselineIncrement, 
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
