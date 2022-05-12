// app.PDFPresetsList.length

var cm = ["GrayScale", "", "RGB", "CMYK"];
preflight();

function preflight() {
    if (app.documents.length > 0) {
        if (ExternalObject.AdobeXMPScript == undefined) {
            ExternalObject.AdobeXMPScript = new ExternalObject('lib:AdobeXMPScript');
        }
        var myDoc = app.documents[0];
        myDoc.rulerOrigin = [0, myDoc.height];
        app.preferences.setIntegerPreference("rulerType", 0);
        var rgb = [], fogra = [], lowres = [];
        for (var i = 0, n = myDoc.placedItems.length; n > i; i++) {
            // var result = [];
            try {
                var placedItem = myDoc.placedItems[i];
                var fNm = placedItem.file;
                var fileName = fNm.name;
                var width = placedItem.width;
                var height = placedItem.height;
                var left = placedItem.left;
                var top = -placedItem.top;

                var xmp = new XMPFile(fNm.fsName, XMPConst.UNKNOWN, XMPConst.OPEN_FOR_READ);
                var obj = xmp.getXMP();
                xmp.closeFile();
                obj.serialize(XMPConst.SERIALIZE_READ_ONLY_PACKET | XMPConst.SERIALIZE_USE_COMPACT_FORMAT);

                var rgbInfo = crossOverRGBImages(obj, left, top, width, height);
                if (rgbInfo) {
                    rgb.push(fileName);
                }
                var fograInfo = isNonFogra39Images(obj);
                if (fograInfo) {
                    fogra.push(fileName);
                }
                var lowresInfo = isLowResImages(obj, width);
                if (lowresInfo != false) {
                    lowres.push(fileName + "\t" + lowresInfo[1] + " actual PPI");
                }
            }
            catch (e) {
                // alert(e);
            }
        }
        var logText = "";
        if (rgb.length > 0 || fogra.length > 0 || lowres.length > 0) {
            if (fogra.length > 0 || lowres.length > 0) {
                if (fogra.length > 0) {
                    logText += "Fogra Information\r=========================================\r"
                    logText += fogra.join("\r");
                    logText += "\r\r";
                }
                if (lowres.length > 0) {
                    logText += "Low-Res Information\r=========================================\r"
                    logText += lowres.join("\r");
                }
                displayLog(myDoc, logText);
            }
        }
        else {
            removeUnusedSwatches();
            setCutterColorToOverprint();
            exportToPDF(myDoc);
        }
    }
    else {
        alert("Open a document and run again.");
    }
}

function crossOverRGBImages(obj, left, top, w, h) {
    try {
        var colorMode = Number(obj.getProperty(XMPConst.NS_PHOTOSHOP, "ColorMode"));
        if (colorMode == 3) {
            crossMark(left, top, w, h);
            return true;
        }
    }
    catch (e) {
        // alert(e);
    }
    return false;
}

function isNonFogra39Images(obj) {
    try {
        var profile = obj.getProperty(XMPConst.NS_PHOTOSHOP, "ICCProfile");
        if (!(/FOGRA/ig.test(profile.toString()))) {
            return true;
        }
    }
    catch (e) {
        // alert(e);
    }
    return false;
}

function isLowResImages(obj, width) {
    try {
        var actualPPI = Number(obj.getProperty(XMPConst.NS_EXIF, "PixelXDimension")) / width * 72;
        if (actualPPI < 250) {
            return [true, Math.round(actualPPI)];
        }
    }
    catch (e) {
        // alert(e);
    }
    return false;
}

function crossMark(left, top, width, height) {
    try {
        var doc = app.activeDocument;
        var group = doc.groupItems.add();
        var right = left + width;
        var bottom = -(top + height);
        var path1 = doc.pathItems.add();
        path1.strokeWidth = 6;
        var p1 = path1.pathPoints.add();
        p1.anchor = p1.rightDirection = p1.leftDirection = [left, -top];
        var p2 = path1.pathPoints.add();
        p2.anchor = p2.rightDirection = p2.leftDirection = [right, bottom];
        path1.moveToEnd(group);

        path1 = doc.pathItems.add();
        path1.strokeWidth = 6;
        p1 = path1.pathPoints.add();
        p1.anchor = p1.rightDirection = p1.leftDirection = [left, bottom];
        p2 = path1.pathPoints.add();
        p2.anchor = p2.rightDirection = p2.leftDirection = [right, -top];
        path1.moveToEnd(group);
    }
    catch (e) {
        // alert(e);
    }
}
function displayLog(myDoc, logText) {
    try {
        var logFile = new File(myDoc.fullName.fsName.replace(/\.ai$/i, "_log.txt"));
        logFile.open('w');
        logFile.write(logText);
        logFile.close();
        logFile.execute();
    }
    catch (e) {
        // alert(e);
    }
}

function exportToPDF(myDoc) {
    try {
        //show a dialog for 
        var preset = getPresetName();
        var options = new PDFSaveOptions();
        options.pDFPreset = preset;
        var pdfFile = new File(myDoc.fullName.fsName.replace(/\.ai$/i, ".pdf"));
        myDoc.saveAs(pdfFile, options);
    }
    catch (e) {
        alert(e);
    }

}

function getPresetName() {

    // DIALOG
    // ======
    var dialog = new Window("dialog");
    dialog.text = "PDF Preset Selection";
    dialog.orientation = "column";
    dialog.alignChildren = ["center", "top"];
    dialog.spacing = 10;
    dialog.margins = 16;

    // PANEL1
    // ======
    var panel1 = dialog.add("panel", undefined, undefined, { name: "panel1" });
    panel1.text = "Preset Selection";
    panel1.orientation = "column";
    panel1.alignChildren = ["left", "top"];
    panel1.spacing = 10;
    panel1.margins = 10;

    // GROUP1
    // ======
    var group1 = panel1.add("group", undefined, { name: "group1" });
    group1.orientation = "row";
    group1.alignChildren = ["left", "center"];
    group1.spacing = 10;
    group1.margins = 0;

    var statictext1 = group1.add("statictext", undefined, undefined, { name: "statictext1" });
    statictext1.text = "Preset: ";

    var dropdown1 = group1.add("dropdownlist", undefined, undefined, { name: "dropdown1", items: app.PDFPresetsList });
    dropdown1.selection = 0;

    // GROUP2
    // ======
    var group2 = dialog.add("group", undefined, { name: "group2" });
    group2.orientation = "row";
    group2.alignChildren = ["left", "center"];
    group2.spacing = 10;
    group2.margins = 0;

    var button1 = group2.add("button", undefined, undefined, { name: "button1" });
    button1.text = "Cancel";
    button1.preferredSize.width = 80;
    button1.preferredSize.height = 30;

    var button2 = group2.add("button", undefined, undefined, { name: "button2" });
    button2.text = "OK";
    button2.preferredSize.width = 80;
    button2.preferredSize.height = 30;

    if (dialog.show() == true) {
        return dropdown1.selection.text;
    }
    else {
        return "";
    }
}

function removeUnusedSwatches() {
    try {
        var actionString = [
            '/version 3',
            '/name [ 8',
            '756e757365647377',
            ']',
            '/isOpen 1',
            '/actionCount 1',
            '/action-1 {',
            '/name [ 8',
            '737774636864656c',
            ']',
            '/keyIndex 0',
            '/colorIndex 0',
            '/isOpen 0',
            '/eventCount 2',
            '/event-1 {',
            '/useRulersIn1stQuadrant 0',
            '/internalName (ai_plugin_swatches)',
            '/localizedName [ 8',
            '5377617463686573',
            ']',
            '/isOpen 0',
            '/isOn 1',
            '/hasDialog 0',
            '/parameterCount 1',
            '/parameter-1 {',
            '/key 1835363957',
            '/showInPalette 4294967295',
            '/type (enumerated)',
            '/name [ 17',
            '    53656c65637420416c6c20556e75736564',
            ']',
            '/value 11',
            '}',
            '}',
            '/event-2 {',
            '/useRulersIn1stQuadrant 0',
            '/internalName (ai_plugin_swatches)',
            '/localizedName [ 8',
            '5377617463686573',
            ']',
            '/isOpen 0',
            '/isOn 1',
            '/hasDialog 1',
            '/showDialog 0',
            '/parameterCount 1',
            '/parameter-1 {',
            '/key 1835363957',
            '/showInPalette 4294967295',
            '/type (enumerated)',
            '/name [ 13',
            '    44656c65746520537761746368',
            ']',
            '/value 3',
            '}',
            '}',
            '}'
        ].join('\n')
        var f = new File(Folder.desktop + "/unusedsw.aia");
        f.open('w');
        f.write(actionString);
        f.close();
        app.loadAction(f);
        app.doScript('swtchdel', 'unusedsw');
        app.unloadAction('unusedsw', '');
        f.remove();
    }
    catch (e) {
        // alert(e);
    }
}

function setCutterColorToOverprint() {
    if (app.documents.length > 0) {
        var doc = app.activeDocument;
        for (var c = 0, cn = doc.swatches.length; cn > c; c++) {
            var color = doc.swatches[c];
            if (color.name == "Cutter") {
                var layers = doc.layers, layer;
                var colorName = color.color.spot.name;
                for (var i = 0, n = layers.length; n > i; i++) {
                    layer = layers[i];
                    if (layer.name == "cutter") {
                        processLayerItems(layer, colorName);
                        break;
                    }
                }
                break;
            }
        }
    }


    function processLayerItems(parent, colorName) {
        try {
            var items = parent.pageItems, item;
            for (var i = 0, n = items.length; n > i; i++) {
                try {
                    item = items[i];
                    // $.writeln(item.typename);
                    if (item.typename == 'PathItem') {
                        try {
                            if (item.filled && item.fillColor.spot.name == colorName) {
                                if (!item.fillOverprint) {
                                    item.fillOverprint = true;
                                }
                            }
                        }
                        catch (e) {
                            // alert(e);
                        }

                        try {
                            if (item.stroked && item.strokeColor.spot.name == colorName) {
                                if (!item.strokeOverprint) {
                                    item.strokeOverprint = true;
                                }
                            }
                        }
                        catch (e) {
                            // alert(e);
                        }
                    }
                    else if (item.typename == 'TextFrame') {
                        try {
                            for (var j = 0, m = item.textRanges.length; m > j; j++) {
                                try {
                                    var textRange = item.textRanges[j];
                                    if (textRange.characterAttributes.fillColor &&
                                        textRange.characterAttributes.fillColor.spot.name == colorName) {
                                        if (!(textRange.characterAttributes.overprintFill)) {
                                            textRange.characterAttributes.overprintFill = true;
                                        }
                                    }
                                }
                                catch (e) {
                                    // alert(e);
                                }
                            };
                        }
                        catch (e) {
                            // alert(e);
                        }

                    }
                    else {
                        processLayerItems(item, colorName);
                    }
                }
                catch (e) {
                    // alert(e);
                }
            }
        }
        catch (e) {
            // alert(e);
        }
    }
}