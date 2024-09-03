// NAME:AutoSubheader
// Status: Working  - This script is constantly being refined - as it currently works on hard coding in attributes which for DT works well but more efficent means could be developed
// FUNCTION - To search for an array of words which are <10 in length not contaning a period (with exceptions) and apply the first instance subheader throughout following text
// The script allows the designer to have many stacked subheaders in which the respective subheader can be applyied throughout the copy 
// OUTLINE - reads the design a user applies to text and applies same design throughout paragraph meeting correct criteria
// This is a sister script to AutoSubheader ---- this script works where there is not bodycopy paragraphs, think of lists such as Product + Price, Gov Spending + £000  etc 
// Prefered keyboard shortcut is Shift + F2


function findAndApplyTextSettings() {
    // Ensure that there's an active document open
    if (app.documents.length > 0) {
        var doc = app.activeDocument;
        var selection = app.selection;

        // Ensure there's a valid selection
        if (selection.length > 0 && (selection[0] instanceof TextFrame || selection[0].hasOwnProperty('insertionPoints'))) {
            var selectedTextFrames = [];
            if (selection[0] instanceof TextFrame) {
                selectedTextFrames.push(selection[0]);
            } else if (selection[0].hasOwnProperty('insertionPoints')) {
                selectedTextFrames.push(selection[0].parentTextFrames[0]);
            }

            var foundTargetParagraph = false;
            var targetTextSettingsList = [];

            // Iterate over each selected text frame
            for (var j = 0; j < selectedTextFrames.length; j++) {
                var paragraphs = selectedTextFrames[j].parentStory.paragraphs.everyItem().getElements();

                // Find the first suitable paragraph and all consecutive ones
                for (var i = 0; i < paragraphs.length; i++) {
                    var consecutiveParagraphs = [];
                    var allConsecutiveSettings = [];

                    for (var k = i; k < paragraphs.length; k++) {
                        var para = paragraphs[k];
                        var words = para.words.length;
                        var paragraphText = para.contents;

                        // Check criteria: 1 to 10 words and no full stop
                        if (words > 0 && words <= 10 && paragraphText.indexOf('.') === -1) {
                            consecutiveParagraphs.push(para);
                            allConsecutiveSettings.push(getTextSettings(para));
                        } else {
                            break; // Stop checking if the paragraph doesn't meet criteria
                        }
                    }

                    if (consecutiveParagraphs.length > 0) {
                        targetTextSettingsList = allConsecutiveSettings;
                        foundTargetParagraph = true;
                        break; // Exit loop once found
                    }
                }

                if (foundTargetParagraph) {
                    // Apply text settings to similar paragraphs within an undoable action
                    app.doScript(function() {
                        var currentConsecutiveIndex = 0;
                        for (var i = 0; i < paragraphs.length; i++) {
                            var para = paragraphs[i];
                            var words = para.words.length;
                            var paragraphText = para.contents;

                            // Apply text settings if it's a similar paragraph
                            if (words > 0 && words <= 10 && paragraphText.indexOf('.') === -1) {
                                if (currentConsecutiveIndex < targetTextSettingsList.length) {
                                    applyTextSettings(para, targetTextSettingsList[currentConsecutiveIndex]);
                                    currentConsecutiveIndex++;
                                }
                            } else {
                                currentConsecutiveIndex = 0; // Reset if paragraph does not meet criteria
                            }
                        }
                    }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Apply Text Settings");
                } else {
                    alert("No suitable paragraphs found in the selected text frame.");
                }
            }
        } else {
            alert("Please select a text frame or place the cursor inside a text frame.");
        }
    } else {
        alert("No active document found.");
    }
}

// Function to get text settings from a paragraph
function getTextSettings(paragraph) {
    return {
        appliedFont: paragraph.appliedFont,
        pointSize: paragraph.pointSize,
        fillColor: paragraph.fillColor,
        fillTint: paragraph.fillTint,
        justification: paragraph.justification,
        leading: paragraph.leading,
        firstLineIndent: paragraph.firstLineIndent,
        baselineShift: paragraph.baselineShift,
        tracking: paragraph.tracking,
        capitalization: getCapitalization(paragraph.capitalization),
        alignToBaseline: paragraph.alignToBaseline,
        ruleAbove: paragraph.ruleAbove === true,
        ruleBelow: paragraph.ruleBelow === true,
        ruleAboveColor: paragraph.ruleAboveColor,
        ruleAboveGapColor: paragraph.ruleAboveGapColor,
        ruleAboveGapOverprint: paragraph.ruleAboveGapOverprint === true,
        ruleAboveGapTint: paragraph.ruleAboveGapTint,
        ruleAboveLeftIndent: paragraph.ruleAboveLeftIndent,
        ruleAboveLineWeight: paragraph.ruleAboveLineWeight,
        ruleAboveOffset: paragraph.ruleAboveOffset,
        ruleAboveOverprint: paragraph.ruleAboveOverprint === true,
        ruleAboveRightIndent: paragraph.ruleAboveRightIndent,
        ruleAboveTint: paragraph.ruleAboveTint,
        ruleAboveType: paragraph.ruleAboveType,
        ruleAboveWidth: paragraph.ruleAboveWidth,
        ruleBelowColor: paragraph.ruleBelowColor,
        ruleBelowGapColor: paragraph.ruleBelowGapColor,
        ruleBelowGapOverprint: paragraph.ruleBelowGapOverprint === true,
        ruleBelowGapTint: paragraph.ruleBelowGapTint,
        ruleBelowLeftIndent: paragraph.ruleBelowLeftIndent,
        ruleBelowLineWeight: paragraph.ruleBelowLineWeight,
        ruleBelowOffset: paragraph.ruleBelowOffset,
        ruleBelowOverprint: paragraph.ruleBelowOverprint === true,
        ruleBelowRightIndent: paragraph.ruleBelowRightIndent,
        ruleBelowTint: paragraph.ruleBelowTint,
        ruleBelowType: paragraph.ruleBelowType,
        ruleBelowWidth: paragraph.ruleBelowWidth,
        spaceBefore: paragraph.spaceBefore, // Add spaceBefore attribute
        spaceAfter: paragraph.spaceAfter // Add spaceAfter attribute
        // Add more text attributes as needed
    };
}

// Function to apply text settings to a paragraph
function applyTextSettings(paragraph, settings) {
    paragraph.appliedFont = settings.appliedFont;
    paragraph.pointSize = settings.pointSize;
    paragraph.fillColor = settings.fillColor;
    paragraph.fillTint = settings.fillTint;
    paragraph.leading = settings.leading;
    paragraph.firstLineIndent = settings.firstLineIndent;
    paragraph.justification = settings.justification;
    paragraph.ruleAbove = settings.ruleAbove;
    paragraph.ruleAboveColor = settings.ruleAboveColor;
    paragraph.ruleAboveGapColor = settings.ruleAboveGapColor;
    paragraph.ruleAboveGapOverprint = settings.ruleAboveGapOverprint;
    paragraph.ruleAboveGapTint = settings.ruleAboveGapTint;
    paragraph.ruleAboveLeftIndent = settings.ruleAboveLeftIndent;
    paragraph.ruleAboveLineWeight = settings.ruleAboveLineWeight;
    paragraph.ruleAboveOffset = settings.ruleAboveOffset;
    paragraph.ruleAboveOverprint = settings.ruleAboveOverprint;
    paragraph.ruleAboveRightIndent = settings.ruleAboveRightIndent;
    paragraph.ruleAboveTint = settings.ruleAboveTint;
    paragraph.ruleAboveType = settings.ruleAboveType;
    paragraph.ruleAboveWidth = settings.ruleAboveWidth;
    paragraph.ruleBelow = settings.ruleBelow;
    paragraph.ruleBelowColor = settings.ruleBelowColor;
    paragraph.ruleBelowGapColor = settings.ruleBelowGapColor;
    paragraph.ruleBelowGapOverprint = settings.ruleBelowGapOverprint;
    paragraph.ruleBelowGapTint = settings.ruleBelowGapTint;
    paragraph.ruleBelowLeftIndent = settings.ruleBelowLeftIndent;
    paragraph.ruleBelowLineWeight = settings.ruleBelowLineWeight;
    paragraph.ruleBelowOffset = settings.ruleBelowOffset;
    paragraph.ruleBelowOverprint = settings.ruleBelowOverprint;
    paragraph.ruleBelowRightIndent = settings.ruleBelowRightIndent;
    paragraph.ruleBelowTint = settings.ruleBelowTint;
    paragraph.ruleBelowType = settings.ruleBelowType;
    paragraph.ruleBelowWidth = settings.ruleBelowWidth;
    paragraph.spaceBefore = settings.spaceBefore; // Apply spaceBefore
    paragraph.spaceAfter = settings.spaceAfter; // Apply spaceAfter

    // Apply word-level settings to each word in the paragraph
    var paraWords = paragraph.words.everyItem().getElements();
    for (var k = 0; k < paraWords.length; k++) {
        paraWords[k].appliedFont = settings.appliedFont;
        paraWords[k].pointSize = settings.pointSize;
        paraWords[k].fillColor = settings.fillColor;
        paraWords[k].fillTint = settings.fillTint;
        paraWords[k].leading = settings.leading;
        paraWords[k].baselineShift = settings.baselineShift;
        paraWords[k].tracking = settings.tracking;
        paraWords[k].capitalization = settings.capitalization;
        paraWords[k].alignToBaseline = settings.alignToBaseline;
        // Apply other text attributes as needed
    }
}

// Helper function to map capitalization constants to InDesign values
function getCapitalization(constantValue) {
    switch (constantValue) {
        case 1634493296:
            return Capitalization.ALL_CAPS;
        case 1664250723:
            return Capitalization.CAP_TO_SMALL_CAP;
        case 1852797549:
            return Capitalization.NORMAL;
        case 1936548720:
            return Capitalization.SMALL_CAPS;
        case 1919251315:
            return true; // For ruleAbove
        case 1919251316:
            return true; // For ruleBelow
        default:
            return Capitalization.NORMAL; // Default to normal capitalization
    }
}

// Run the function
findAndApplyTextSettings();
