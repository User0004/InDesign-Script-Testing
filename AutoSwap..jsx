// NAME:AutoSwap
// Status: Working 
// FUNCTION - To preform a swap of two selections.
// Selections can be of mixed media i.e. image swap with text box and so on. 
// OUTLINE - Works by preforming the same set of events a user would undertake ensuring that database items are not mixed up
// Picture crops are defaulted to Options.FILL_PROPORTIONALLY - an updated version of the script tried to prevent crop changes for selections made where images were the exact same size - although tolerance would be nice as users may not have exact same size
// One can select two single objects or  - group an array of objects together with a selection to move wider selections - one can also use this script to swap stories on the page with new furniture and not have to redraw an entire page
// Prefered keyboard shortcut is F11



app.doScript(function () {
    if (app.selection.length !== 2) {
        alert("Please select exactly two objects.");
    } else {
        var firstObject = app.selection[0];
        var secondObject = app.selection[1];

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
                return null; // Return null if the object does not contain a text frame
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
                        adjustImagePosition(group.allPageItems[i], offsetX, offsetY);
                    }
                }
            }

            // Function to adjust image position within the new frame
            function adjustImagePosition(object, offsetX, offsetY) {
                if (object.constructor.name === "Rectangle" && object.images.length > 0) {
                    var image = object.images[0];
                    image.geometricBounds = [
                        image.geometricBounds[0] + offsetY,
                        image.geometricBounds[1] + offsetX,
                        image.geometricBounds[2] + offsetY,
                        image.geometricBounds[3] + offsetX
                    ];
                }
            }

            // Check if the dimensions (width and height) are equal
            if (areDimensionsEqual(firstWidth, firstHeight, secondWidth, secondHeight)) {
                // Calculate the offset between the first and second object
                var offsetX = secondBounds[1] - firstBounds[1];
                var offsetY = secondBounds[0] - firstBounds[0];

                // Adjust the image position for both objects
                adjustImagePosition(firstObject, offsetX, offsetY);
                adjustImagePosition(secondObject, -offsetX, -offsetY);

                // Adjust image positions within groups if necessary
                adjustGroupImagesPosition(firstObject, offsetX, offsetY);
                adjustGroupImagesPosition(secondObject, -offsetX, -offsetY);

            } else {
                // Function to fit content in images
                function fitContentInImages(object, fitOption) {
                    if (object.constructor.name === "Rectangle" && object.images.length > 0) {
                        object.fit(fitOption);
                    } else if (object.constructor.name === "Group") {
                        for (var i = 0; i < object.allPageItems.length; i++) {
                            if (object.allPageItems[i].constructor.name === "Rectangle" && object.allPageItems[i].images.length > 0) {
                                object.allPageItems[i].fit(fitOption);
                            }
                        }
                    }
                }

                // Use FILL_PROPORTIONALLY option for fitting images
                var fitOptions = FitOptions.FILL_PROPORTIONALLY;

                // Fit content in images in both objects using FILL_PROPORTIONALLY
                fitContentInImages(firstObject, fitOptions);
                fitContentInImages(secondObject, fitOptions);

                // Adjust image positions within groups if necessary
                adjustGroupImagesPosition(firstObject, 0, 0);
                adjustGroupImagesPosition(secondObject, 0, 0);
            }
        }
    }
}, ScriptLanguage.JAVASCRIPT, null, UndoModes.ENTIRE_SCRIPT, "Swap Object Sizes and Positions");
