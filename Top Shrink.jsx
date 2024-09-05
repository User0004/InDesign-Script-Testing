// NAME:Top Expand , Top Shrink, Bottom Expand, Bottom Shrink 
// Status: Working 
// FUNCTION - To perform an increase or decrease of the height of one or more selected items - which vary in their heights - by the unit of one baseline for each time a script is run 
// InDesign cannot by default allow a user to change the height of all boxes in a selection at the same time while maintaining a fixed position of selections opposite to that being expanded or shrinked 
// The script means a designer can expand or shrink an entire story in one go 
// OUTLINE - Script reads baseline grid size and either adds or subtracks baseline height to overall height of selections 

// Prefered keyboard shortcut is Top Expand F9, Top Shrink Shift+9, Bottom Expand F10 , Bottom Shrink Shift F10





//Top Shrink

// Define the script as a function to be executed
app.doScript(function() {
    // Get the active document
    var doc = app.activeDocument;

    // Get document grid preferences
    var gridPreferences = doc.gridPreferences;
    var baselineIncrement = gridPreferences.baselineDivision;

    // Function to adjust height from the top downwards
    function adjustHeight(item) {
        var currentHeight = item.geometricBounds[2] - item.geometricBounds[0];
        
        if (currentHeight <= baselineIncrement) {
            // If the height is less than or equal to the baseline increment, move the item downwards
            var direction = (baselineIncrement > 0) ? 1 : -1; // Text flows down or up
            
            // Adjust geometric bounds to move downwards by one baseline increment
            item.geometricBounds = [
                item.geometricBounds[0] + direction * baselineIncrement,  // Top edge Y-coordinate moves downwards
                item.geometricBounds[1],  // Left edge X-coordinate remains the same
                item.geometricBounds[2] + direction * baselineIncrement, // Bottom edge moves downwards
                item.geometricBounds[3]  // Right edge X-coordinate remains the same
            ];
        } else {
            // If the height is greater than the baseline increment, adjust normally
            item.geometricBounds = [
                item.geometricBounds[0] + baselineIncrement, // Adjust top edge downwards
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
