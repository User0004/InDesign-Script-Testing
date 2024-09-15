// NAME: AutoSubheader
// Status: Ready
// FUNCTION: To search for an array of words which are under 10 in length and apply the first instance subheader throughout following text
// The script allows the designer to have many stacked subheaders that can be applied throughout the copy with as many enters intween as they like
// This is a sister script AutoList 
// Preferred keyboard shortcut is F2

// AutoSubheader
// Function to get the index of the insertion point in a text frame
function getInsertionPoint(textFrame) {
    var selection = app.selection;
    var cursorIndex = -1;

    if (selection.length > 0) {
        for (var i = 0; i < selection.length; i++) {
            if (selection[i] instanceof Text) {
                cursorIndex = selection[i].index;
                break;
            } else if (selection[i] instanceof InsertionPoint) {
                cursorIndex = selection[i].index;
                break;
            }
        }
    }

    if (cursorIndex === -1) {
        // Default to 0 if no valid cursor position found
        cursorIndex = 0;
    }

    return cursorIndex;
}

// Function to process a single text frame
function processTextFrame(textFrame) {
    var paragraphs = textFrame.parentStory.paragraphs.everyItem().getElements();
    var cursorIndex = getInsertionPoint(textFrame);

    var foundTargetParagraph = false;
    var targetTextSettingsList = [];
    var applyIndentAdjustment = false;

    var consecutiveCount = 0;
    var consecutiveParagraphs = [];
    var allConsecutiveSettings = [];
    var lastValidParagraph = null;

    for (var i = 0; i < paragraphs.length; i++) {
        var paragraph = paragraphs[i];

        // Skip paragraphs before the insertion point
        if (paragraph.index < cursorIndex) {
            continue;
        }

        if (paragraph.contents.replace(/^\s+|\s+$/g, '').length === 0) {
            continue;
        }

        if (isValidParagraph(paragraph)) {
            consecutiveCount++;
            consecutiveParagraphs.push(paragraph);
            allConsecutiveSettings.push(getTextSettings(paragraph));
            lastValidParagraph = paragraph;

            applyIndentAdjustment = true;
        } else {
            if (consecutiveCount > 0) {
                if (!foundTargetParagraph) {
                    targetTextSettingsList = allConsecutiveSettings;
                    foundTargetParagraph = true;
                }
                applyTextSettingsToConsecutive(consecutiveParagraphs, targetTextSettingsList);
                consecutiveCount = 0;
                consecutiveParagraphs = [];
                allConsecutiveSettings = [];
                lastValidParagraph = null;
                applyIndentAdjustment = false;
            } else {
                consecutiveCount = 0;
                consecutiveParagraphs = [];
                allConsecutiveSettings = [];
                lastValidParagraph = null;
                applyIndentAdjustment = false;
            }
        }

        if (applyIndentAdjustment) {
            for (var k = i + 1; k < paragraphs.length; k++) {
                var nextParagraph = paragraphs[k];
                if (nextParagraph.contents.replace(/^\s+|\s+$/g, '').length > 0) {
                    nextParagraph.firstLineIndent = 0;
                    applyIndentAdjustment = false;
                    break;
                }
            }
        }
    }

    if (foundTargetParagraph) {
        // Handle cases where settings were found
    } else {
        // Handle cases where no subheaders were found
    }
}

// Function to check if a paragraph meets the criteria
function isValidParagraph(paragraph) {
    var contents = paragraph.contents.replace(/^\s+|\s+$/g, '');
    
    if (contents.length === 0) {
        return false;
    }

    var words = paragraph.words;
    if (words.length > 0 && words.length <= 10) {
        var lastWord = words.lastItem();
        var lastWordContents = lastWord.contents;

        // Check if the last word ends with a period
        if (lastWordContents.charAt(lastWordContents.length - 1) === ".") {
            // Exclude paragraphs that end with a period
            return false;
        }

        return true;
    }

    return false;
}

// Function to get text settings from a paragraph
function getTextSettings(paragraph) {
    var settings = {
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
        spaceBefore: paragraph.spaceBefore,
        spaceAfter: paragraph.spaceAfter,
        hyphenation: paragraph.hyphenation
    };

    var paragraphStyle = paragraph.appliedParagraphStyle;
    if (paragraphStyle.isValid && isStyleUnmodified(paragraph, paragraphStyle)) {
        settings.paragraphStyle = paragraphStyle;
    } else {
        settings.paragraphStyle = null;
    }

    return settings;
}

// Function to check if the paragraph matches its base paragraph style
function isStyleUnmodified(paragraph, paragraphStyle) {
    var styleProperties = ['appliedFont', 'pointSize', 'fillColor', 'fillTint', 'justification', 'leading', 'firstLineIndent', 'baselineShift', 'tracking', 'capitalization', 'alignToBaseline', 'ruleAbove', 'ruleBelow', 'spaceBefore', 'spaceAfter'];

    for (var i = 0; i < styleProperties.length; i++) {
        var prop = styleProperties[i];
        if (paragraph[prop] !== paragraphStyle[prop]) {
            return false;
        }
    }
    return true;
}

// Function to apply text settings to a paragraph
function applyTextSettings(paragraph, settings) {
    if (settings.paragraphStyle) {
        paragraph.applyParagraphStyle(settings.paragraphStyle, true);
    } else {
        paragraph.appliedFont = settings.appliedFont;
        paragraph.pointSize = settings.pointSize;
        paragraph.fillColor = settings.fillColor;
        paragraph.fillTint = settings.fillTint;
        paragraph.leading = settings.leading;
        paragraph.firstLineIndent = settings.firstLineIndent;
        paragraph.justification = settings.justification;
        paragraph.baselineShift = settings.baselineShift;
        paragraph.tracking = settings.tracking;
        paragraph.capitalization = settings.capitalization;
        paragraph.alignToBaseline = settings.alignToBaseline;
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
        paragraph.spaceBefore = settings.spaceBefore;
        paragraph.spaceAfter = settings.spaceAfter;
        paragraph.hyphenation = settings.hyphenation;
    }
}

// Function to apply text settings to consecutive paragraphs
function applyTextSettingsToConsecutive(paragraphs, settingsList) {
    app.doScript(function() {
        var currentConsecutiveIndex = 0;
        for (var i = 0; i < paragraphs.length; i++) {
            var para = paragraphs[i];
            var words = para.words.length;
            var paragraphText = para.contents;

            if (words > 0 && words <= 10 && !/^\d+(\.\d+)?$/.test(paragraphText)) {
                if (currentConsecutiveIndex < settingsList.length) {
                    applyTextSettings(para, settingsList[currentConsecutiveIndex]);
                    currentConsecutiveIndex++;
                }
            } else {
                currentConsecutiveIndex = 0;
            }
        }
    }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Apply Text Settings");
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
            return true;
        case 1919251316:
            return true;
        default:
            return Capitalization.NORMAL;
    }
}

// Main function to find and apply text settings to consecutive paragraphs
function findAndApplyTextSettings() {
    if (app.documents.length > 0) {
        var doc = app.activeDocument;
        var selection = app.selection;

        if (selection.length > 0) {
            var selectedTextFrames = [];

            for (var i = 0; i < selection.length; i++) {
                if (selection[i] instanceof TextFrame) {
                    selectedTextFrames.push(selection[i]);
                } else if (selection[i] instanceof Text) {
                    selectedTextFrames.push(selection[i].parentTextFrames[0]);
                } else if (selection[i] instanceof InsertionPoint) {
                    selectedTextFrames.push(selection[i].parentTextFrames[0]);
                }
            }

            if (selectedTextFrames.length === 0) {
                alert("No valid text frames found in selection.");
                return;
            }

            for (var j = 0; j < selectedTextFrames.length; j++) {
                processTextFrame(selectedTextFrames[j]);
            }
        } else {
            alert("Please select at least one text frame or insertion point.");
        }
    } else {
        alert("No active document found.");
    }
}

// Run the function to find and apply text settings
app.doScript(function() {
    findAndApplyTextSettings();
}, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Find and Apply Text Settings");
