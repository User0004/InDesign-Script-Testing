// NAME:AutoSwap
// STATUS: Working 
// FUNCTION: To preform a swap of two selections.
// Selections can be of mixed media i.e. image swap with text box and so on. 
// OUTLINE: Works by preforming the same set of events a user would undertake ensuring that database items are not mixed up
// Picture crops are defaulted to Options.FILL_PROPORTIONALLY - an updated version of the script tried to prevent crop changes for selections made where images were the exact same size - although tolerance would be nice as users may not have exact same size
// One can select two single objects or  - group an array of objects together with a selection to move wider selections - one can also use this script to swap stories on the page with new furniture and not have to redraw an entire page
// Prefered keyboard shortcut is F11
// Updated code to handle swaps of differnt SplineItem subclasses, namely Oval | Rectangle | Polygon
// Note user will be alerted to an error if they try to swap something with a GraphicLine -- aka just a straight line 



//AutoSwap
//AutoSwap

app.doScript(function () {
    if (app.selection.length !== 2) {
        alert("Please select exactly two objects.");
    } else {
        var firstObject = app.selection[0];
        var secondObject = app.selection[1];

        // Check if either object is a GraphicLine
        if (firstObject.constructor.name === "GraphicLine" || secondObject.constructor.name === "GraphicLine") {
            alert("Cannot swap selection with a rule.");
            return; // Exit the script if a GraphicLine is detected
        }

        // Ensure both objects are valid
        if (!firstObject.isValid || !secondObject.isValid) {
            alert("Invalid selection. Please select two valid objects.");
        } else {
            // Function to remember text frame settings
            function rememberTextFrameSettings(object) {
                if (object.constructor.name === "TextFrame") {
                    return {
                        verticalJustification: object.textFramePreferences.verticalJustification,
                        columnCount: object.textFramePreferences.textColumnCount,
                        isMultiColumn: object.textFramePreferences.textColumnCount > 1
                    };
                }
                return null;
            }

            // Function to apply text frame settings
            function applyTextFrameSettings(object, settings) {
                if (settings !== null) {
                    object.textFramePreferences.verticalJustification = settings.verticalJustification;
                    object.textFramePreferences.textColumnCount = settings.columnCount;
                }
            }

            // Remember text frame settings
            var firstSettings = rememberTextFrameSettings(firstObject);
            var secondSettings = rememberTextFrameSettings(secondObject);

            // Get the geometric bounds of both objects
            var firstBounds = firstObject.geometricBounds;
            var secondBounds = secondObject.geometricBounds;

            // Calculate width and height from geometric bounds
            var firstWidth = firstBounds[3] - firstBounds[1];
            var firstHeight = firstBounds[2] - firstBounds[0];
            var secondWidth = secondBounds[3] - secondBounds[1];
            var secondHeight = secondBounds[2] - secondBounds[0];

            // Function to compare width and height
            function areDimensionsEqual(width1, height1, width2, height2) {
                return Math.abs(width1 - width2) < 6 && Math.abs(height1 - height2) < 6; // Tolerance for floating-point precision
            }

            // Check if either object is a multi-column text frame
            var isFirstMultiColumn = firstSettings && firstSettings.isMultiColumn;
            var isSecondMultiColumn = secondSettings && secondSettings.isMultiColumn;

            // Handle swapping based on the type of objects
            if ((isFirstMultiColumn && !isSecondMultiColumn && secondObject.constructor.name === "Rectangle") ||
                (!isFirstMultiColumn && isSecondMultiColumn && firstObject.constructor.name === "Rectangle")) {
                // Swap geometric bounds only for multi-column text frame with picture
                firstObject.geometricBounds = secondBounds;
                secondObject.geometricBounds = firstBounds;
            } else {
                // Swap geometric bounds and apply text frame settings
                firstObject.geometricBounds = secondBounds;
                secondObject.geometricBounds = firstBounds;

                // Apply text frame settings only if both are text frames
                if (firstSettings && secondSettings) {
                    applyTextFrameSettings(firstObject, secondSettings);
                    applyTextFrameSettings(secondObject, firstSettings);
                }
            }

            // Handle grouped objects specifically
            function adjustGroupImagesPosition(group, offsetX, offsetY) {
                if (group.constructor.name === "Group") {
                    for (var i = 0; i < group.allPageItems.length; i++) {
                        adjustMediaPosition(group.allPageItems[i], offsetX, offsetY);
                    }
                }
            }

            // Function to adjust media position (images and graphics) within the new frame
            function adjustMediaPosition(object, offsetX, offsetY) {
                if ((object.constructor.name === "Rectangle" || 
                     object.constructor.name === "Oval" || 
                     object.constructor.name === "Polygon") && 
                    (object.images.length > 0 || object.graphics.length > 0)) {
                    
                    // Handle both images and graphics (including .eps)
                    var media = object.images.length > 0 ? object.images[0] : object.graphics[0];

                    // Adjust geometric bounds for all shapes
                    media.geometricBounds = [
                        media.geometricBounds[0] + offsetY,
                        media.geometricBounds[1] + offsetX,
                        media.geometricBounds[2] + offsetY,
                        media.geometricBounds[3] + offsetX
                    ];
                }
            }

            // Function to fit content in images/graphics with proportional adjustment for all shapes
            function fitContentInMediaProportionally(object, fitOption) {
                if ((object.constructor.name === "Rectangle" || 
                     object.constructor.name === "Oval" || 
                     object.constructor.name === "Polygon") && 
                    (object.images.length > 0 || object.graphics.length > 0)) {

                    object.fit(fitOption);  // Use proportional fitting
                } else if (object.constructor.name === "Group") {
                    for (var i = 0; i < object.allPageItems.length; i++) {
                        fitContentInMediaProportionally(object.allPageItems[i], fitOption);
                    }
                }
            }

            // Check if the dimensions (width and height) are equal
            if (areDimensionsEqual(firstWidth, firstHeight, secondWidth, secondHeight)) {
                // Calculate the offset between the first and second object
                var offsetX = secondBounds[1] - firstBounds[1];
                var offsetY = secondBounds[0] - firstBounds[0];

                // Adjust the image/graphic position for both objects
                adjustMediaPosition(firstObject, offsetX, offsetY);
                adjustMediaPosition(secondObject, -offsetX, -offsetY);

                // Adjust media positions within groups if necessary
                adjustGroupImagesPosition(firstObject, offsetX, offsetY);
                adjustGroupImagesPosition(secondObject, -offsetX, -offsetY);

            } else {
                // Use FILL_PROPORTIONALLY option for fitting media without stretching
                var fitOptions = FitOptions.FILL_PROPORTIONALLY;

                // Fit content in media in both objects using proportional fitting for all shapes
                fitContentInMediaProportionally(firstObject, fitOptions);
                fitContentInMediaProportionally(secondObject, fitOptions);

                // Adjust media positions within groups if necessary
                adjustGroupImagesPosition(firstObject, 0, 0);
                adjustGroupImagesPosition(secondObject, 0, 0);
            }
        }
    }
}, ScriptLanguage.JAVASCRIPT, null, UndoModes.ENTIRE_SCRIPT, "Swap Object Sizes and Positions");
