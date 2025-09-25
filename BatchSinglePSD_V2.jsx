// ######################################
// 
// 
// Project: Photoshop Automation Plugin for a single PSD file and multiple Images to Enbed/Insert
// Date: 2023-05-23 
// Author: Muyideen .A. Jimoh
// Github: https://github.com/JIMOHMA
// 
// ######################################
var FOUNDLAYER = undefined;

var activeDoc = app.activeDocument;
var activeDocName = activeDoc.name;
$.writeln(activeDocName.length);
var activeDoclayerSets = activeDoc.layerSets;
editPlaceMultipleArtWork() // Entry Point

// function to check if running on a Mac machine
function isMacOS() {
    return $.os.indexOf("Mac") !== -1
}

// function to check if runnign on a Windows machine
function isWindowsOS() {
    return $.os.indexOf("Windows") !== -1
}

function editPlaceMultipleArtWork() {
    alert("Select the artWork Folder...📸🎈🎈🎨🎨📸");
    var folderARTWORK = getfolderPATH();  // folderARTWORK is the path/address to the folder containing the images we want to embedd
    // get all the files in a folder
    var containingFiles = folderARTWORK.getFiles();

    // Filter out only the PNG and JPG files
    const filteredFiles = [];
    for (var ind = 0; ind < containingFiles.length; ind++) {
        $.writeln('containingFiles[ind] is ', containingFiles[ind].fsName);
        if ((containingFiles[ind].displayName.indexOf('.png') != -1) || (containingFiles[ind].displayName.indexOf('.jpg') != -1)){
            filteredFiles.push(containingFiles[ind].displayName);
        }
    }

    // enbedd each JPG or PNG file into the PSD file one after the other and save
    for (var index = 0; index < filteredFiles.length; index++) {
        var artName = filteredFiles[index];

        activeDocument = app.documents[0];
        var myDoc = activeDocument;
        if (FOUNDLAYER != undefined) {
            try {
                processFoundLayer(myDoc, artName, folderARTWORK); // skip the recursive look for the layer of interest and use the found layer.
            } catch (error) {
                alert("Error is: " + error.message + "\n"+ "On line: "+ error.line)
            }
        } else {
            goThroughLayers(myDoc, artName, folderARTWORK); // inside this goThroughLayers() function 
        }

                                                        // is a recursive call to look through nested layer
    }
}

function processFoundLayer(parentLayer, artName, folderARTWORK) {
    openSmartObject(FOUNDLAYER);  // opens the layer of interest in a new tab or document
                        
    // now we're ready to add new ArtWork and save the file
    var artDesignFileName = artName;
    activeDocument.selection.selectAll(); // CTRL + A 
    // first fill background with white color and save document
    fillWithWhite();
    // rename layer to "YOUR DESIGN HERE" if artist has used a different name
    // alert("Renaming smart Layer where artWork is placed");
    try {
        if (activeDocument.layers[0].name != "YOUR DESIGN HERE") {
            activeDocument.layers[0].name = "YOUR DESIGN HERE";
        } 
    } catch (error) {
        alert("Document contains no layer!");
    }
    activeDocument.save();
    // plase artwork of interest
    if (isMacOS()) {
        placeImageOnMac(artDesignFileName, folderARTWORK);
    } 
    if (isWindowsOS()) {
        placeArtWork(artDesignFileName, folderARTWORK);
        // placeImage(imageNameWithExt, imagesFolder);
    }

    // Option 1: Centre active layer on canvas
    // alignCenterOfCanvas(); 

    // Option 2: Fill the entire canvas with the attached image 
    transformToOriginalLayerPosition();
    activeDocument.save(); // this allows the image attached to appear in the original PSD file within the smart object layer
    
    // then save a copy of the newly created EDIT using "Save A Copy" as JPEG option. 
    // then change the active document to the first document and save file
    activeDocument = app.documents[0];
    try {
        var doc = app.activeDocument;
        jpgSaveOptions = new JPEGSaveOptions();
        jpgSaveOptions.embedColorProfile = true;
        jpgSaveOptions.formatOptions = FormatOptions.STANDARDBASELINE;
        jpgSaveOptions.matte = MatteType.NONE;
        jpgSaveOptions.quality = 10;

        var Path = doc.path;  
        $.writeln("Location of temp saved file is " + Path); 
        var parentDocumentName = doc.name.replace(/\.[^\.]+$/, ''); 
        var Suffix = "-Copy";
        try {
            var artWithNoJpg = artDesignFileName.replace('.png','');
        }
        catch (e) {
            var artWithNoJpg = artDesignFileName.replace('.jpg','');
        }

        try {
            var saveFolder = new Folder(Path + "/"+ parentDocumentName); // folder ovject with path/location where it will reside
            if (!saveFolder.exists) {
                saveFolder.create();
                var saveFile = File(saveFolder.fullName + "/"+ parentDocumentName + "-" + artWithNoJpg + Suffix );
                $.writeln("Save lcnif: ", saveFile);
                activeDocument.saveAs(saveFile, jpgSaveOptions, true, Extension.LOWERCASE);
            }
            else {
                var saveFile = File(Path + "/" + parentDocumentName + "/"+ parentDocumentName + "-" + artWithNoJpg + Suffix );
                $.writeln("Save lcnelse: ", saveFile);
                activeDocument.saveAs(saveFile, jpgSaveOptions, true, Extension.LOWERCASE);
            }
        }
        catch (e) {
            alert("Something went wrong...Unable to create/find folder and save new Edit.....");
        }
    }
    catch(e) {
        alert("Document has not been saved yet!");
    }

    // make the newly opened document the active document 
    // This is usualy the 2nd document but this needs revision 
    // in order to logically not run into issues when users use the application
    // So we need to dynamically know which document we're working with....
    // Perhaps knowing the names of the document and referencing them based on their 
    // names rather than some arbitrary collection indexing??? 
    // TODO: Find an implementation for this.
    activeDocument = app.documents[1];

    // 1. Delete every layer that's not called "YOUR DESIGN HERE"
    // 2. Save the mockup
    // 3. Switch back to 1st document i.e. make active and Continue the iteration 
    var condition = ""; // condition can either be "Delete", "", or "Can't Delete"
    while (condition != "Can't Delete") {
        condition = deletePlacedLayers(); //step 1
    }
    activeDocument.save(); // step 2
    activeDocument = app.documents[0]; // step 3
}

// Function to get the PATH to a selected folder from a dialog box
function getfolderPATH() {
    var myFolder = Folder.selectDialog("Selecting a folder where your artWorks are saved..... 😂😎😆🦾💻🙄😇🤑");
    return myFolder
}

// deletes a placed layer
function deletePlacedLayers() {
    var layer = app.activeDocument.activeLayer;
    var layerName = layer.name;

    if (layerName != "YOUR DESIGN HERE") {
        // activeDocument.save();
        layer.remove(); // Delete placed layer
        return "Deleted"; 
    }
    else {
        // alert("Can't delete 'YOUR DESIGN HERE' layer");
        return "Can't Delete";
    }
}

// go through the layers and edit the smart object layer
function goThroughLayers(parentLayer, artName, folderARTWORK){

    for(var i=0;i<parentLayer.layers.length;i++){
        var curLayer = parentLayer.layers[i];

        try {
            // activeDoc.activeLayer = curLayer;  // This line makes every layer visible which I don't want
            if(curLayer.typename =='LayerSet'){
                // alert("Found layer ", curLayer.name);
                goThroughLayers(curLayer, artName, folderARTWORK); // recursive function call to itself
                
            } // end if
            else{
                if((curLayer.name.indexOf('double click') != -1) || objectIsPsObject(curLayer)){
                    activeDoc.activeLayer = curLayer;
                    if (objectIsPsObject(activeDoc.activeLayer)) {
                        FOUNDLAYER = curLayer
                        openSmartObject(curLayer);  // opens the layer of interest in a new tab or document
                        
                        // now we're ready to add new ArtWork and save the file
                        var artDesignFileName = artName;
                        activeDocument.selection.selectAll(); // CTRL + A 
                        // first fill background with white color and save document
                        fillWithWhite();
                        // rename layer to "YOUR DESIGN HERE" if artist has used a different name
                        // alert("Renaming smart Layer where artWork is placed");
                        try {
                            if (activeDocument.layers[0].name != "YOUR DESIGN HERE") {
                                activeDocument.layers[0].name = "YOUR DESIGN HERE";
                            } 
                        } catch (error) {
                            alert("Document contains no layer!");
                        }
                        activeDocument.save();
                        // plase artwork of interest
                        if (isMacOS()) {
                            placeImageOnMac(artDesignFileName, folderARTWORK);
                        } 
                        if (isWindowsOS()) {
                            placeArtWork(artDesignFileName, folderARTWORK);
                            // placeImage(imageNameWithExt, imagesFolder);
                        }

                        // Option 1: Centre active layer on canvas
                        // alignCenterOfCanvas(); 

                        // Option 2: Fill the entire canvas with the attached image 
                        transformToOriginalLayerPosition();
                        activeDocument.save(); // this allows the image attached to appear in the original PSD file within the smart object layer
                        
                        // then save a copy of the newly created EDIT using "Save A Copy" as JPEG option. 
                        // then change the active document to the first document and save file
                        activeDocument = app.documents[0];
                        try {
                            var doc = app.activeDocument;
                            jpgSaveOptions = new JPEGSaveOptions();
                            jpgSaveOptions.embedColorProfile = true;
                            jpgSaveOptions.formatOptions = FormatOptions.STANDARDBASELINE;
                            jpgSaveOptions.matte = MatteType.NONE;
                            jpgSaveOptions.quality = 10;
        
                            var Path = doc.path;  
                            $.writeln("Location of temp saved file is " + Path); 
                            var parentDocumentName = doc.name.replace(/\.[^\.]+$/, ''); 
                            var Suffix = "-Copy";
                            try {
                                var artWithNoJpg = artDesignFileName.replace('.png','');
                            }
                            catch (e) {
                                var artWithNoJpg = artDesignFileName.replace('.jpg','');
                            }

                            try {
                                var saveFolder = new Folder(Path + "/"+ parentDocumentName); // folder ovject with path/location where it will reside
                                if (!saveFolder.exists) {
                                    saveFolder.create();
                                    var saveFile = File(saveFolder.fullName + "/"+ parentDocumentName + "-" + artWithNoJpg + Suffix );
                                    $.writeln("Save lcnif: ", saveFile);
                                    activeDocument.saveAs(saveFile, jpgSaveOptions, true, Extension.LOWERCASE);
                                }
                                else {
                                    var saveFile = File(Path + "/" + parentDocumentName + "/"+ parentDocumentName + "-" + artWithNoJpg + Suffix );
                                    $.writeln("Save lcnelse: ", saveFile);
                                    activeDocument.saveAs(saveFile, jpgSaveOptions, true, Extension.LOWERCASE);
                                }
                            }
                            catch (e) {
                                alert("Something went wrong...Unable to create/find folder and save new Edit.....");
                            }
                        }
                        catch(e) {
                            alert("Document has not been saved yet!");
                        }

                        // make the newly opened document the active document 
                        // This is usualy the 2nd document but this needs revision 
                        // in order to logically not run into issues when users use the application
                        // So we need to dynamically know which document we're working with....
                        // Perhaps knowing the names of the document and referencing them based on their 
                        // names rather than some arbitrary collection indexing??? 
                        // TODO: Find an implementation for this.
                        activeDocument = app.documents[1];

                        // 1. Delete every layer that's not called "YOUR DESIGN HERE"
                        // 2. Save the mockup
                        // 3. Switch back to 1st document i.e. make active and Continue the iteration 
                        var condition = ""; // condition can either be "Delete", "", or "Can't Delete"
                        while (condition != "Can't Delete") {
                            condition = deletePlacedLayers(); //step 1
                        }
                        activeDocument.save(); // step 2
                        activeDocument = app.documents[0]; // step 3
                    }
                    return; 
                }
            }
        } catch (error) {
            alert("My error is ", error);
        }
    }
}

// Calls the FitLayerToCasvas function which does the translation and resizing
function transformToOriginalLayerPosition() {
    // FIT LAYER TO CANVAS
    // Referenced via https://forums.adobe.com/message/5413957#5413957
    var maintainAspectRatio;
    if(app.documents.length>0){  
        app.activeDocument.suspendHistory ('Fit Layer to Canvas', 'FitLayerToCanvas('+maintainAspectRatio+')');  
    }  
}

// Reusable function to translate and resize to fill an image to the entire canvas
function FitLayerToCanvas( keepAspect ){// keepAspect:Boolean - optional. Default to false  
    var doc = app.activeDocument;  
    var layer = doc.activeLayer;  
    // do nothing if layer is background or locked  
    if(layer.isBackgroundLayer || layer.allLocked || layer.pixelsLocked  
                            || layer.positionLocked || layer.transparentPixelsLocked ) return;  
    // do nothing if layer is not normal artLayer or Smart Object  
    if( layer.kind != LayerKind.NORMAL && layer.kind != LayerKind.SMARTOBJECT) return;  
    // store the ruler  
    var defaultRulerUnits = app.preferences.rulerUnits;  
    app.preferences.rulerUnits = Units.PIXELS;  
    
    var width = doc.width.as('px');  
    var height =doc.height.as('px');  
    var bounds = app.activeDocument.activeLayer.bounds;  
    var layerWidth = bounds[2].as('px')-bounds[0].as('px');  
    var layerHeight = bounds[3].as('px')-bounds[1].as('px');  
        
    // move the layer so top left corner matches canvas top left corner  
    layer.translate(new UnitValue(0-layer.bounds[0].as('px'),'px'), new UnitValue(0-layer.bounds[1].as('px'),'px'));  
    if( !keepAspect ){  
        // scale the layer to match canvas  
        layer.resize( (width/layerWidth)*100,(height/layerHeight)*100,AnchorPosition.TOPLEFT);  
    }else{  
        var layerRatio = layerWidth / layerHeight;  
        var newWidth = width;  
        var newHeight = ((1.0 * width) / layerRatio);  
        if (newHeight >= height) {  
            newWidth = layerRatio * height;  
            newHeight = height;  
        }  
        var resizePercent = newWidth/layerWidth*100;  
        app.activeDocument.activeLayer.resize(resizePercent,resizePercent,AnchorPosition.TOPLEFT);  
    }  
    // restore the ruler  
    app.preferences.rulerUnits = defaultRulerUnits;  
}

// Aligns the active document to the centre of the Canvas
function alignCenterOfCanvas() {
    var doc = app.activeDocument;
    var docLay = app.activeDocument.activeLayer;
    docLay.translate(doc.width / 2 - (docLay.bounds[0] + docLay.bounds[2]) / 2, 
    doc.height / 2 - (docLay.bounds[1] + docLay.bounds[3]) / 2);
}

/**
 * Place an image by name and location/path where it resides <= macOS implementation
 * @param {*} fileName 
 * @param {*} folderPath 
 */
function placeImageOnMac(fileName, folderPath) {
    // =======================================================
    var idplaceEvent = stringIDToTypeID( "placeEvent" );
        var desc263 = new ActionDescriptor();
        var idID = stringIDToTypeID( "ID" );
        desc263.putInteger( idID, 1715 );
        var idnull = stringIDToTypeID( "null" );
        desc263.putPath( idnull, new File( folderPath + '/' + fileName ) );
        var idfreeTransformCenterState = stringIDToTypeID( "freeTransformCenterState" );
        var idquadCenterState = stringIDToTypeID( "quadCenterState" );
        var idQCSAverage = stringIDToTypeID( "QCSAverage" );
        desc263.putEnumerated( idfreeTransformCenterState, idquadCenterState, idQCSAverage );
        var idoffset = stringIDToTypeID( "offset" );
            var desc264 = new ActionDescriptor();
            var idhorizontal = stringIDToTypeID( "horizontal" );
            var idpixelsUnit = stringIDToTypeID( "pixelsUnit" );
            desc264.putUnitDouble( idhorizontal, idpixelsUnit, 0.000000 );
            var idvertical = stringIDToTypeID( "vertical" );
            var idpixelsUnit = stringIDToTypeID( "pixelsUnit" );
            desc264.putUnitDouble( idvertical, idpixelsUnit, 0.000000 );
        var idoffset = stringIDToTypeID( "offset" );
        desc263.putObject( idoffset, idoffset, desc264 );
        var idreplaceLayer = stringIDToTypeID( "replaceLayer" );
            var desc265 = new ActionDescriptor();
            var idfrom = stringIDToTypeID( "from" );
                var ref11 = new ActionReference();
                var idlayer = stringIDToTypeID( "layer" );
                ref11.putIdentifier( idlayer, 1714 );
            desc265.putReference( idfrom, ref11 );
            var idto = stringIDToTypeID( "to" );
                var ref12 = new ActionReference();
                var idlayer = stringIDToTypeID( "layer" );
                ref12.putIdentifier( idlayer, 1715 );
            desc265.putReference( idto, ref12 );
        var idplaceEvent = stringIDToTypeID( "placeEvent" );
        desc263.putObject( idreplaceLayer, idplaceEvent, desc265 );
    executeAction( idplaceEvent, desc263, DialogModes.NO );
}

// custom function to place an artWork (a .png or .jpg file)
function placeArtWork(fileName, folderARTWORK) {
    
    var idPlc = charIDToTypeID( "Plc " );
        var desc1353 = new ActionDescriptor();
        var idIdnt = charIDToTypeID( "Idnt" );
        desc1353.putInteger( idIdnt, 309 );
        var idnull = charIDToTypeID( "null" );
        
        // coded such that I can add any file I want
        desc1353.putPath( idnull, new File( folderARTWORK + '\\' + fileName) );
        var idFTcs = charIDToTypeID( "FTcs" );
        var idQCSt = charIDToTypeID( "QCSt" );
        var idQcsa = charIDToTypeID( "Qcsa" );
        desc1353.putEnumerated( idFTcs, idQCSt, idQcsa );
        
        
        var idOfst = charIDToTypeID( "Ofst" );
        var desc1354 = new ActionDescriptor();
        var idHrzn = charIDToTypeID( "Hrzn" );
        var idRlt = charIDToTypeID( "#Rlt" );
        desc1354.putUnitDouble( idHrzn, idRlt, -8.263100 );
        var idVrtc = charIDToTypeID( "Vrtc" );
        var idRlt = charIDToTypeID( "#Rlt" );
        desc1354.putUnitDouble( idVrtc, idRlt, 0.000000 );
        var idOfst = charIDToTypeID( "Ofst" );
        desc1353.putObject( idOfst, idOfst, desc1354 );
    executeAction( idPlc, desc1353, DialogModes.NO );
}

// waits for photoshop to open the active layer in a new document/tab successfully
function waitForRedraw() {
    var eventWait = charIDToTypeID("Wait");
    var enumRedrawComplete = charIDToTypeID("RdCm");
    var typeState = charIDToTypeID("Stte");
    var keyState = charIDToTypeID("Stte");
    var desc = new ActionDescriptor();

    desc.putEnumerated(keyState, typeState, enumRedrawComplete);
    executeAction(eventWait, desc, DialogModes.NO);
    return;
}

// function to open a smart object
function openSmartObject(theLayer) {
    current = activeDoc.activeLayer;
    
    if (theLayer.kind == "LayerKind.SMARTOBJECT"){
        //this runMenuItem is indicating that a menu was selected, and in this case it was the "Place Embedded"
        // menu that was selected
        app.runMenuItem(stringIDToTypeID('placedLayerEditContents')); 
        waitForRedraw();  // waits for photoshop to open layer in a new tab
    }
}

// lazy way for filling in what using the scriptlistener
function fillWithWhite() {
    try {
        var idFl = charIDToTypeID( "Fl  " );
        var desc345 = new ActionDescriptor();
        var idUsng = charIDToTypeID( "Usng" );
        var idFlCn = charIDToTypeID( "FlCn" );
        var idClr = charIDToTypeID( "Clr " );
        desc345.putEnumerated( idUsng, idFlCn, idClr );
        var idClr = charIDToTypeID( "Clr " );
            var desc346 = new ActionDescriptor();
            var idH = charIDToTypeID( "H   " );
            var idAng = charIDToTypeID( "#Ang" );
            desc346.putUnitDouble( idH, idAng, 117.174683 );
            var idStrt = charIDToTypeID( "Strt" );
            desc346.putDouble( idStrt, 0.000000 );
            var idBrgh = charIDToTypeID( "Brgh" );
            desc346.putDouble( idBrgh, 100.000000 );
        var idHSBC = charIDToTypeID( "HSBC" );
        desc345.putObject( idClr, idHSBC, desc346 );
        var idOpct = charIDToTypeID( "Opct" );
        var idPrc = charIDToTypeID( "#Prc" );
        desc345.putUnitDouble( idOpct, idPrc, 100.000000 );
        var idMd = charIDToTypeID( "Md  " );
        var idBlnM = charIDToTypeID( "BlnM" );
        var idNrml = charIDToTypeID( "Nrml" );
        desc345.putEnumerated( idMd, idBlnM, idNrml );
        executeAction( idFl, desc345, DialogModes.NO );
        // alert("Completed white fill");
    }
    catch (e) {
        // alert("Solid color, white, and merge");
        // =======================================================
        var idMk = charIDToTypeID( "Mk  " );
        var desc351 = new ActionDescriptor();
        var idnull = charIDToTypeID( "null" );
        var ref31 = new ActionReference();
        var idcontentLayer = stringIDToTypeID( "contentLayer" );
        ref31.putClass( idcontentLayer );
        desc351.putReference( idnull, ref31 );
        var idUsng = charIDToTypeID( "Usng" );
        var desc352 = new ActionDescriptor();
        var idType = charIDToTypeID( "Type" );
        var desc353 = new ActionDescriptor();
        var idClr = charIDToTypeID( "Clr " );
        var desc354 = new ActionDescriptor();
        var idRd = charIDToTypeID( "Rd  " );
        desc354.putDouble( idRd, 255.000000 );
        var idGrn = charIDToTypeID( "Grn " );
        desc354.putDouble( idGrn, 255.000000 );
        var idBl = charIDToTypeID( "Bl  " );
        desc354.putDouble( idBl, 255.000000 );
        var idRGBC = charIDToTypeID( "RGBC" );
        desc353.putObject( idClr, idRGBC, desc354 );
        var idsolidColorLayer = stringIDToTypeID( "solidColorLayer" );
        desc352.putObject( idType, idsolidColorLayer, desc353 );
        var idcontentLayer = stringIDToTypeID( "contentLayer" );
        desc351.putObject( idUsng, idcontentLayer, desc352 );
        executeAction( idMk, desc351, DialogModes.NO );
        // =======================================================
        var idMrgV = charIDToTypeID( "MrgV" );
        executeAction( idMrgV, undefined, DialogModes.NO );
        // alert("Done Catch Fill");
    }   
}

// checks if a layer can be opened in photoshop or not. 
// returns True if it can and False if it can't (meaning opens with another imaging application)
function objectIsPsObject(SOlayer) {
	//Thanks to r-bin
	var ext = smartobject_file_ext(SOlayer);  
    var rc = true;
	switch (ext)  
		{  
		case "nef":  
		case "cr2":  
		case "crw":  
		case "raf":  
		case "orf":  
		case "mrw":  
		case "dcr":  
		case "mos":  
		case "raw":  
		case "pef":  
		case "srf":  
		case "dng":  
		case "x3f":  
		case "erf":  
		case "sr2":  
		case "kdc":  
		case "mfw":  
		case "mef":  
		case "arw":  
		case "nrw":  
		case "rw2":  
		case "rwl":  
		case "iiq":  
		case "3fr":  
		case "fff":  
		case "srw":  
		case "ai":
		case "svg":
		case "pdf":
		case "esp":
			rc = false.	
			break;  
		case "error":  
			rc = false;
			break;         
  		default:  
			rc = true;
			break;  
    }  
	return rc;
}

// function to obtain the extension of a layer 
function smartobject_file_ext(layer) {  
    try {         
        var r = new ActionReference();     
        r.putProperty(stringIDToTypeID("property"), stringIDToTypeID("smartObject"));  
        r.putIdentifier(stringIDToTypeID("layer"), layer.id);  
        var name = executeActionGet(r).getObjectValue(stringIDToTypeID("smartObject")).getString(stringIDToTypeID("fileReference"));         
         
        var n = name.lastIndexOf(".");  
        if (n < 0) return "";  
     
        return name.substr(n+1).toLowerCase();  
        }  
    catch (e) { return "error"; }  
}

alert("Batch is done....🤝🤝🙏🤝🤝");