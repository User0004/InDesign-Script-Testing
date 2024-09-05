// NAME:AutoRules
// Status: Working 
// FUNCTION - Adds rules inbetween selections of differnt types i.e. graphic frames text boxes etc 
// OUTLINE - Pairs with another script AutoRules Controller (below) which creates a json file containing the attribues of the rules wanted by a designer 
// If a designer does not need to run the AutoRules Controller one may simply run the AutoRules which as been set up to work from default settings appropriate for DT pages 
// The script places a 0.3pt black rule with 
// The script can work out how high or low to place rules based on neighbouring heights of selections PLUS selections can be made across spreads without placing rules in the gutter 
// Prefered keyboard shortcut is F6




// AutoRule script 
// Set measurement units to points
app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;

// Load user-defined values from XML
var f = File($.fileName).parent + "/autoRule_user_preset.xml";
var xml = readXML(f);

// Check if xml is empty, and if so, provide default XML
if (xml == "") {
    xml = XML("<dialog><values><top>-3</top><bottom>0</bottom><stroke>0.3</stroke><color>3</color></values></dialog>");
}

var userDefinedAmountTop = Number(xml.values.top);
var userDefinedAmountBottom = Number(xml.values.bottom);
var userDefinedStrokeWeight = Number(xml.values.stroke);
var userDefinedColorIndex = Number(xml.values.color); 

// Function to check if all selected frames are on a page and not on the pasteboard
function areFramesOnPage(frames) {
    for (var i = 0; i < frames.length; i++) {
        if (!frames[i].parentPage) {
            return false; // At least one frame is on the pasteboard
        }
    }
    return true; // All frames are on the page
}

// Main code below
if (app.documents.length > 0) {
    var doc = app.activeDocument;
    var selectedFrames = doc.selection;

    if (!areFramesOnPage(selectedFrames)) {
        alert("Please ensure that all selected frames are on the page, not on the pasteboard.");
    } else if (selectedFrames.length >= 2) {
        selectedFrames.sort(function(a, b) {
            return a.geometricBounds[1] - b.geometricBounds[1];
        });

        app.doScript(function() {
            // Get the color swatch based on the user-defined index
            var selectedColor = doc.swatches[userDefinedColorIndex];

            for (var i = 0; i < selectedFrames.length - 1; i++) {
                var currentFrame = selectedFrames[i];
                var nextFrame = selectedFrames[i + 1];

                if (currentFrame.parentPage === nextFrame.parentPage) {
                    var centerX = (currentFrame.geometricBounds[3] + nextFrame.geometricBounds[1]) / 2;
                    var ruleHeight = Math.max(
                        currentFrame.geometricBounds[2] - currentFrame.geometricBounds[0],
                        nextFrame.geometricBounds[2] - nextFrame.geometricBounds[0]
                    );

                    var startHeight = Math.max(
                        currentFrame.geometricBounds[2],
                        nextFrame.geometricBounds[2]
                    );

                    var endHeight = Math.min(
                        currentFrame.geometricBounds[0],
                        nextFrame.geometricBounds[0]
                    );

                    var adjustedStartHeight = startHeight + userDefinedAmountBottom;
                    var adjustedEndHeight = endHeight - userDefinedAmountTop;

                    var newLine = currentFrame.parentPage.graphicLines.add();
                    newLine.strokeWeight = userDefinedStrokeWeight;
                    newLine.paths[0].entirePath = [
                        [centerX, adjustedStartHeight],
                        [centerX, adjustedEndHeight]
                    ];

                    // Apply the selected color to the stroke
                    newLine.strokeColor = selectedColor;

                    newLine.sendToBack();
                }
            }
        }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Add Vertical Lines Between Text Boxes");
    } else {
        alert("Please select at least two text frames.");
    }
} else {
    alert("Open a document before running this script.");
}

// Reset measurement units to auto value
app.scriptPreferences.measurementUnit = AutoEnum.AUTO_VALUE;

function readXML(p) {
    var f = new File(p);
    f.open("r");
    var t = f.read();
    f.close();
    return XML(t);
}
