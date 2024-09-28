// NAME:AutoRules Controller 
// STATUS: Ready
// FUNCTION - Adds rules inbetween selections of differnt types i.e. graphic frames text boxes etc 
// OUTLINE - Pairs with another script AutoRules Controller (below) which creates a xml file containing the attribues of the rules wanted by a designer 
// If a designer does not need to run the AutoRules Controller one may simply run the AutoRules which as been set up to work from default settings appropriate for DT pages 
// The script has some preset values which are the default values making it easy back to the regular rules
// The script can work out how high or low to place rules based on neighbouring heights of selections PLUS selections can be made across spreads without placing rules in the gutter 
// Author: George Hannaford george.hannaford@telegraph.co.uk
// Prefered keyboard not needed as it is rare that it should be needed but well worth keeping for special supplements etc 
// Version 1.0.0



// AutoRules Controller
var f = File($.fileName).parent + "/autoRule_user_preset.xml";

var xml = readXML(f);
// Check to see if xml file is equal to nothing, if not present use below
if (xml == "") {
    xml = XML("<dialog><values><top>2</top><bottom>1</bottom><stroke>1</stroke><color>0</color></values></dialog>");
}

var tv = Number(xml.values.top);  // var for topValue
var bv = Number(xml.values.bottom);  // var for bottomValue
var sv = Number(xml.values.stroke);  // var for strokeValue
var cv = Number(xml.values.color);  // var for colourValue

var t, b, s, c, cd;
makeDialog();

// Creates the box for user input
function makeDialog() {
    var d = app.dialogs.add({ name: "autoRules Controller", canCancel: true });
    with (d.dialogColumns.add()) {
        staticTexts.add({ staticLabel: "Adjust top height:" });
        staticTexts.add({ staticLabel: "Adjust bottom height:" });
        staticTexts.add({ staticLabel: "Adjust stroke weight:" });
        staticTexts.add({ staticLabel: "Stroke colour:" });
    }
    with (d.dialogColumns.add()) {
        // Set fixed default values directly
        t = measurementEditboxes.add({ editUnits: MeasurementUnits.POINTS, editValue: -3, minWidth: 50 }); // var for topValue
        b = measurementEditboxes.add({ editUnits: MeasurementUnits.POINTS, editValue: 0, minWidth: 50 }); // var for bottomValue
        s = measurementEditboxes.add({ editUnits: MeasurementUnits.POINTS, editValue: 0.3, minWidth: 50 });//  var for strokeValue
        cd = dropdowns.add({ stringList: getArrayNames(app.activeDocument.swatches.everyItem().getElements()), selectedIndex: cv || 0, minWidth: 80 });
    }
    if (d.show() == true) {
        t = t.editValue;
        b = b.editValue;
        s = s.editValue;
        // get the swatch object
        c = app.activeDocument.swatches.everyItem().getElements()[cd.selectedIndex];

        // update the XML with new values (optional)
        xml.values.top = t;
        xml.values.bottom = b;
        xml.values.stroke = s;
        xml.values.color = cd.selectedIndex; // the selected dropdown number
        writeXML(f, xml.toXMLString());

        // call the main() function via doScript here
        app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'Add Vertical Lines Between Text Boxes');
        d.destroy();
    }
}

/**
 * Returns an array of swatch names for the dropdown 
 * @return {Array}
 */
function getArrayNames(arr) {
    var a = [];
    for (var i = 0; i < arr.length; i++) {
        a.push(arr[i].name);
    }
    return a;
}

/**
 * Checks if all selected frames are on a page and not on the pasteboard.
 * @param {Array} frames - The array of selected frames.
 * @return {Boolean} - True if all frames are on a page, false if any frame is on the pasteboard.
 */
function areFramesOnPage(frames) {
    for (var i = 0; i < frames.length; i++) {
        if (!frames[i].parentPage) {
            return false; // At least one frame is on the pasteboard
        }
    }
    return true; // All frames are on the page
}

/**
 * The main script 
 * The main script may use points but it should be following parameters set via the controller so long as user sets their unit in the input box then script will convert said unit 
 */
function main() {
    // the dialog is getting points, so make sure the script uses points no matter what the ruler units are set to
    app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;

    if (app.documents.length > 0) {
        var doc = app.activeDocument;
        var selectedFrames = doc.selection;

        if (!areFramesOnPage(selectedFrames)) {
            alert("Please ensure that all selected frames are on the page, not on the pasteboard.");
        } else if (selectedFrames.length >= 2) {
            selectedFrames.sort(function (a, b) {
                return a.geometricBounds[1] - b.geometricBounds[1];
            });

            for (var i = 0; i < selectedFrames.length - 1; i++) {
                var currentFrame = selectedFrames[i];
                var nextFrame = selectedFrames[i + 1];

                if (currentFrame.parentPage === nextFrame.parentPage) {
                    var centerX = (currentFrame.geometricBounds[3] + nextFrame.geometricBounds[1]) / 2;
                    var startHeight = Math.max(currentFrame.geometricBounds[2], nextFrame.geometricBounds[2]);
                    var endHeight = Math.min(currentFrame.geometricBounds[0], nextFrame.geometricBounds[0]);
                    var adjustedStartHeight = startHeight + b;
                    var adjustedEndHeight = endHeight - t;

                    // graphic lines have a bounds not including the stroke
                    var newLine = currentFrame.parentPage.graphicLines.add({
                        geometricBounds: [adjustedEndHeight, centerX, adjustedStartHeight, centerX],
                        strokeWeight: s,
                        strokeColor: c
                    });

                    newLine.sendToBack();
                }
            }
        } else {
            alert("Please select at least two text frames.");
        }
    } else {
        alert("Open a document before running this script.");
    }

    app.scriptPreferences.measurementUnit = AutoEnum.AUTO_VALUE;
}

function writeXML(p, s) {
    var file = new File(p);
    file.encoding = 'UTF-8';
    file.open('w');
    file.write(s);
    file.close();
}

function readXML(p) {
    var f = new File(p);
    f.open("r");
    var t = f.read();
    f.close();
    return XML(t);
}
