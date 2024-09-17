/*
# Analysis of yeast cells in dMSCC chips
## Luca Torello Pianale (Chalmers University, lucat@chalmers.se) 
## Luisa Blöbaum (Bielefeld University, luisa.bloebaum@uni-bielefeld.de)

__Macro Properly Working on ImageJ 1.54j.__
__July 2024__

------------------

This macro is used to _analyse yeast cells in microfluidic chips_ (fluorescence, area, shape descriptors, etc.). 

__General info__:
* No need of heavy pre-processing of the phase contrast stacks to get the ROIs, the StarDist 2D model (Yeast_Detection_v2.2) has been trained to recognise directly the phase contrast images of yeast cells.
* Run the macro in FIJI (www.fiji.sc) either dragging the .ijm in FIJI or selecting the file from `Plugins / Macros / Run...`.
* The language can be both IJM Macro or IJM Macro Markdown, if a final HTML file is desired.
* __Plugins Needed__: Bio-Formats, Template matching, StarDist 2D, ResultsToExcel, Masks from ROIs, IJMMD (if final HTML of the macro is required). If not installed already, install them from `Help / Update... / Manage update sites`. 
* For plugin Template_matching: use the built-in update manager (ImageJ's Menu>Help>Update...), and add the following URL via "Manage update sites">"Add update site" -> http://sites.imagej.net/Template_Matching/ 

The __Input__ folder should contain .nd2 or .ome files for each chamber in the microfluidic chip that one wants to analyse. 
The file name should contain AT LEAST an identifier for each chamber at the end. E.g.: "FileName_XY000.nd2", where 000 stands for the number of the chamber. You can select at the beginning which your identifier is (XY in this case).

The __Macro_Image_Analysis__ folder has this macro saved in it and the subfolder Yeast_Detection_v2.2 in which the model used for StarDis2d is. The model can be stored also sowhere else!

The __Output__ folder should contain subfolders named:
* _"Results"_ (to save the information associated to each ROI as an .xslx file, as well as two additional files summarising the data for easier plotting outside of R). 
* _"ROIs_StarDist2D"_ (to save the ROIs coming from StarDist2D for each chamber as .zip files). 
* _"ROIs_TrackMate"_ (to save the ROIs coming from TrackMate for each chamber as .zip files). 
* _"Hyperstacks"_ (to save the hyperstacks for each chamber as .tif files). 
* _"Hyperstacks_fluo"_ (to save the fluorescence ratio hyperstacks for each chamber as .tif files). 
* _"Edges"_ (to save the edges coming from TrackMAte for each chamber as .csv files).
* _"XML"_ (to save the TrackMate file for each chamber as .xml files).

If these are not the input/output formats, changes should be applied in the macro. 

Troubleshooting-Guide:
* Sometimes, thresholding for cutting does not work. The macro will then save the uncut version. These positions need to be cut manually.
* ome.tiff files large than 2 GB will not load. Use single nd files for each position instead.
* While opening, the files are not opend correctly. Make sure that nothing in the Bio-Formats import mask is ticked. 

*/
/*
--------------------
# FRESH START FOR FIJI.

*/

close("*");
roiManager("reset");
run("Clear Results");
print("\\Clear");
roiManager("Show None");
roiManager("Associate", "true");
roiManager("Centered", "false");
roiManager("UseNames", "false");

/*
------------------
# SET FOLDER PATHS.

*/

Dialog.create("Before Starting!");
Dialog.addMessage("Make in the 'Output' folder the following sub-folders:\n-Results\n-ROIs_StarDist2D\n-ROIs_TrackMate\n-Hyperstacks\n-Hyperstacks_fluo\n-Edges\n-XML");
Dialog.addMessage("If ready, you can go on!");
Dialog.show();

input = getDirectory("Load files from directory...");
output = getDirectory("Save processed files to directory...");

/*
-------------------
# EXPERIMENT DETAILS.

Set all the experimental details for proper analysis.

*/

strainchoice = newArray("Choose!", "CENPK", "EthanolRed", "PE2");
sensorchoice = newArray("Choose!", "No Biosensor", "QUEEN", "sfpHluorin", "GlyOx", "RibUPR", "PyruEth", "Other");
channelchoice = newArray(" ", "phase_contrast", "uvgfp", "cfp", "gfp", "yfp", "rfp");
colorchoice = newArray(" ", "Grays", "Blue", "Cyan", "Green", "Yellow", "Red");
PDMSchoice = newArray("top", "left", "bottom");
YesNo = newArray("Choose!", "Yes", "No");
ROIchoice = newArray("Choose!", "StarDist2D", "TrackMate");

Dialog.create("Experimental presets");
	Dialog.addNumber("Date (YYMMDD)", 230223);
	Dialog.addChoice("Strain", strainchoice);
	Dialog.addChoice("Sensor", sensorchoice);
	Dialog.addNumber("Time interval (min between pictures)", 8);
	
	Dialog.addNumber("Number of channels", 1);
	Dialog.addNumber("Stack position of phase contrast channel", 1);
	
	Dialog.addChoice("Reference PDMS", PDMSchoice, "top");
	Dialog.addString("Chamber identifier (e.g., XY)", "XY");
	
	Dialog.addChoice("Make Hyperstacks?", YesNo);
	Dialog.addChoice("Run StarDist2D?", YesNo);
	Dialog.addChoice("Run TrackMate?", YesNo);
	Dialog.addChoice("Run data analysis?", YesNo);

	Dialog.addMessage("Name your channels and choose a display color");
	Dialog.addChoice("Channel 1", channelchoice);
	Dialog.addToSameRow();
	Dialog.addChoice("Color", colorchoice);
	Dialog.addChoice("Channel 2", channelchoice);
	Dialog.addToSameRow();
	Dialog.addChoice("Color", colorchoice);
	Dialog.addChoice("Channel 3", channelchoice);
	Dialog.addToSameRow();
	Dialog.addChoice("Color", colorchoice);
	Dialog.addChoice("Channel 4", channelchoice);
	Dialog.addToSameRow();
	Dialog.addChoice("Color", colorchoice);
Dialog.show();

date = Dialog.getNumber(); 
strain = Dialog.getChoice();
sensor = Dialog.getChoice();
interval = Dialog.getNumber(); 

num_channels = Dialog.getNumber();
phase_channel = Dialog.getNumber();

PDMSwhere = Dialog.getChoice();
chamberidentifier = Dialog.getString();

makehyperstacks = Dialog.getChoice();
runstardist = Dialog.getChoice();
runtrackmate = Dialog.getChoice();
runanalysis = Dialog.getChoice();

ch1 = Dialog.getChoice();
col1 = Dialog.getChoice();
ch2 = Dialog.getChoice();
col2 = Dialog.getChoice();
ch3 = Dialog.getChoice();
col3 = Dialog.getChoice();
ch4 = Dialog.getChoice();
col4 = Dialog.getChoice();

color = newArray(col1, col2, col3, col4);
pos = newArray(1, 2, 3, 4);
channelname = newArray(ch1, ch2, ch3, ch4);

color = Array.slice(color, 0, num_channels);
pos = Array.slice(pos, 0, num_channels);
channelname = Array.slice(channelname, 0, num_channels);
channeltitle = String.join(channelname);

print("Date (YYMMDD): " + date + "\nS. cerevisiae Strain: " + strain + "\nBiosensor: " + sensor + "\nPicture Time Inteval (min): " + interval + "\nChannel Order: " + channeltitle)
selectWindow("Log");
saveAs("Text", output + "Experimental_Details#1.txt");

/*
--------------------------------------------
# Making HYPERSTACKS.

Importing the raw data and making hyperstacks for easier analysis later.

*/

if (makehyperstacks == "Yes") {

	list_folders = getFileList(input);
	
	for (i = 0; i < list_folders.length; i++){
		print("Making hyperstack: " + i + 1 + " of " + list_folders.length);
	
		//Open
	 	run("Bio-Formats Importer", "open=[" + input + list_folders[i] + "] autoscale color_mode=Grayscale rois_import=[ROI manager] view=Hyperstack stack_order=XYCZT");
	 	run("Stack to Hyperstack...", "order=xyczt(default) channels=" + num_channels + " slices=1 frames=" + nSlices/num_channels + " display=Grayscale");
		run("Set Scale...", "distance=0 known=0 unit=pixel"); //No scale should be applied for now
		rename("C");
		
		if (PDMSwhere == "left") run("Rotate 90 Degrees Left");
		if (PDMSwhere == "bottom") run("Rotate... ", "angle=180 grid=1 interpolation=Bilinear stack");
	
		//Setting chamber rotation angle and threshold.
		if (i == 0) {
			beep();
			setTool("line"); 
			waitForUser("Draw rotation line", "Draw a line along the light refractory on the upper chamber edge. Then press 'OK'.");
			run("Set Measurements...", "mean standard min stack display redirect=None decimal=3");
			run("Measure");
			angle = getResult("Angle", 0);
			leng = round(getResult("Length", 0));
			getLine(x1, y1, x2, y2, lineWidth);
			print("Angle: " + angle);
			getSelectionCoordinates(xpoints, ypoints);
			p = Array.findMinima(ypoints, 0);
			miny = ypoints[p[0]];
			
			run("Select None");
			run("Clear Results");
			} 
			
		getDimensions(width, height, channels, slices, frames); //only needed for width and height
		
		//Align phase contrast stack.
		//X and Y positions might need to be changed upon need.
		selectWindow("C");
		if (num_channels > 1) run("Split Channels");
		if (num_channels == 1) rename("C1-C");
		
		selectWindow("C" + phase_channel + "-C"); 
		run("Align slices in stack...", "method=5 windowsizex=" + leng + " windowsizey=" + miny-60 + " x0=" + x1 + " y0=35 swindow=100 subpixel=true itpmethod=0 ref.slice=1 show=true");	
		run("Select None");
		run("Duplicate...", "title=find_chamber duplicate range=1-1");
		run("Rotate... ", "angle=" + angle + " grid=0 interpolation=Bilinear");
		
		//Align fluorescent stacks.
		for (k = 1; k <= num_channels; k++) { 
			if (k != phase_channel) {
				selectWindow("C" + k + "-C");
				for (chani = 2; chani <= nSlices; chani++) {
				    setSlice(chani);
				    chanihelp = chani-2;
				    x = getResult("dX", chanihelp);
				    y = getResult("dY", chanihelp);
				    run("Translate...", "x=" + x + " y=" + y + " interpolation=None");
				}}}
		
		//Make hyperstack.
		if (num_channels == 2) run("Merge Channels...", "c1=[C" + pos[0] + "-C] c2=[C" + pos[1] +"-C] create");
		if (num_channels == 3) run("Merge Channels...", "c1=[C" + pos[0] + "-C] c2=[C" + pos[1] + "-C] c3=[C" + pos[2] +"-C] create");
		if (num_channels == 4) run("Merge Channels...", "c1=[C" + pos[0] + "-C] c2=[C" + pos[1] + "-C] c3=[C" + pos[2] +"-C] c4=[C" + pos[3] +"-C] create");
		
		//Rotate (hyper)stack.	
		if (num_channels > 1) selectWindow("C");
		if (num_channels == 1) selectWindow("C1-C");
		run("Rotate... ", "angle=" + angle + " grid=0 interpolation=Bilinear stack");
		
		//Find Threshold for the chamber outline.	
		selectWindow("find_chamber");
		run("Enhance Contrast...", "saturated=0.90");
		run("Smooth");
		run("Invert");
	
		// find min and max of the grayvalues for thresholding along the line perpendicular line
		
		if (PDMSwhere == "top" || PDMSwhere == "bottom"){
			mid = (x2 - x1)/2+x1;
			makeLine(mid, y1+100, mid, y2-100);	
		}
		
		if (PDMSwhere == "left"){
			mid = (y2 - y1)/2+y1;
			makeLine(x1-100, mid, x2+100, mid);	
		}
		
		run("Measure");
		minthresh = getResult("Min", 0);
		maxthresh = getResult("Max", 0);
		run("Clear Results");
		
		//set the half of those values as lower threshold, upper is always max
		lowerthresh = minthresh + (maxthresh-minthresh)/2;
		setAutoThreshold("Default dark");
		run("Threshold...");
		setThreshold(lowerthresh, 65535);
		run("Convert to Mask");
	
		if (isOpen("Threshold")) {selectWindow("Threshold"); run("Close");};
	
		//Detect chamber
		selectWindow("find_chamber");
		run("Analyze Particles...", "size=800000-Infinity exclude overlay add");
		if (roiManager("count") != 0) {
			roiManager("Select", 0);
			run("Enlarge...", "enlarge=-15");
			run("Fit Rectangle");
			
			//Crop (hyper)stack.
			roiManager("Add");
			if (num_channels != 1) {selectWindow("C");}
			if (num_channels == 1) {selectWindow("C1-C");}
			roiManager("Select", 1);
			run("Crop");
			run("Select None");
		}
		selectWindow("find_chamber");
		close();
		roiManager("reset");
		
		//Process images (background subtraction, enhancement, etc..).
		//For background subtraction: the rolling ball radius might vary based on the instrument/microorganism.
		
		if (num_channels > 1) {selectWindow("C"); run("Split Channels");} 
	
		for (k = 1; k <= num_channels; k++) {
			if (k != phase_channel) {
				selectWindow("C" + k + "-C");
				run("Subtract Background...", "rolling=150 stack"); 
				run("32-bit");
				setAutoThreshold("Default dark");
				run("NaN Background", "stack");
				rename("C" + k);		
			} else {
				selectWindow("C" + k + "-C");
				rename("C" + k);
				run("32-bit");
				run("Sharpen", "stack");
				run("Despeckle", "stack");	
				}}
		
		//Make final hyperstack. 
		//NOTE: it is possible also to just convert all the channels to gray for simplicity.
		time = nSlices;
		
		if (num_channels == 2) run("Merge Channels...", "c1=[C" + pos[0] + "] c2=[C" + pos[1] +"] create");
		if (num_channels == 3) run("Merge Channels...", "c1=[C" + pos[0] + "] c2=[C" + pos[1] + "] c3=[C" + pos[2] +"] create");
		if (num_channels == 4) run("Merge Channels...", "c1=[C" + pos[0] + "] c2=[C" + pos[1] + "] c3=[C" + pos[2] +"] c4=[C" + pos[3] +"] create");
	
		run("Properties...", "channels=" + num_channels +" slices=1 frames=" + time);
		
		//Adjust LUTs.
		for (k = 1; k <= num_channels; k++) {
			Stack.setChannel(k);
			run(color[k-1]);
			}
	
		//Add details to each stack (change scale parameters upon need).
		run("Set Scale...", "distance=13.9 known=1 unit=um");
		setForegroundColor(252, 252, 252);
		run("Scale Bar...", "width=10 height=10 font=28 color=White background=Black location=[Upper Right] bold overlay label");
	    
		//Save hyperstack.
		//Note: the name of the original file should finish with "XY000.nd2", where "000" stands for the number of the chamber.
	 	chamber = substring(list_folders[i], indexOf(list_folders[i], chamberidentifier), lastIndexOf(list_folders[i], ".nd2")); 
		rename(date + "_" + strain + "_" + sensor + "_Channels " + channeltitle + "_" + chamber);
		name = getTitle();
		saveAs("Tiff", output + "Hyperstacks/" + name + ".tif");
		
		close("*");
		print("\\Clear");
		run("Clear Results");
		beep();
	}
}

/*
-------------------------------------------------
# STARDIST2D ANALYSIS.

This automated loop uses __StarDist2D__ and a trained model to identify the cells from phase-contrast images. 
Fluorescent images will then be analysed, if needed, in a later step. 

For __StarDist2D__ to work: 
* Select the folder where the "TF_SavedModel.zip" file is (possibly in the "Macro_Image_Analysis" folder).
* Settings in StarDist2D can be changed directly in the code upon need.

*/

if (runstardist == "Yes") { 
	
	modelfolder = replace(getDirectory("Select StarDis2D Model Folder..."), "\\", "/");
	
	list_stacks = getFileList(output + "Hyperstacks/");
	
	for (i = 0; i < list_stacks.length; i++){
		print("Analysing stack " + i + 1 + " of " + list_stacks.length);
		
		//Open Hyperstack.
		filename = output + "Hyperstacks/" + list_stacks[i];
		open(filename);
		rename(replace(getTitle(), ".tif", ""));
		name = getTitle();
		chamber = substring(name, indexOf(name, chamberidentifier), lastIndexOf(name, "")); 
	
		//Split Channels.
		run("Duplicate...", "duplicate");
		rename("temporary");
		selectWindow(name);
		close();
		selectWindow("temporary");
		if (num_channels == 1) rename("C1");
		if (num_channels > 1) {
			run("Split Channels");
			for (k = 1; k <= num_channels; k++){	
				selectWindow("C" + k + "-temporary");
				rename("C" + k);
			}
		}
				
		//Run StardDist2D.
		selectWindow("C" + phase_channel);
		print("Analysing Chamber: " + chamber);
		
		run("Command From Macro", "command=[de.csbdresden.stardist.StarDist2D], args=['input':'C" + phase_channel + "', 'modelChoice':'Model (.zip) from File', 'normalizeInput':'true', 'percentileBottom':'1.5', 'percentileTop':'99.5', 'probThresh':'0.6', 'nmsThresh':'0.4', 'outputType':'ROI Manager', 'modelFile':'" + modelfolder + "TF_SavedModel.zip', 'nTiles':'1', 'excludeBoundary':'40', 'roiPosition':'Automatic', 'verbose':'false', 'showCsbdeepProgress':'false', 'showProbAndDist':'false'], process=[false]");
		
		//Save the ROIs	
		roiManager("Save", output + "ROIs_StarDist2D/RoiSet_" + name + ".zip");
	   
		//Prepare "Fresh Start" for next loop.
		roiManager("reset");
		run("Clear Results");
		print("\\Clear");
		close("*");
	}
}

/*
-------------------------------------------------
# TRACKMATE ANALYSIS.

This loop uses TrackMate to allow tracking of individual cells and identification of budding events (establishing mother-daugther relationships).  
This section of the macro is semi-automated, meaning that still requires inputs from the user for each file uploaded. 
We suggest to run first all the stacks and perform the analysis (next section) immidiately, then edit the lineages afterwards. 

In __TrackMate__: 
* As detector use "Label image detector" if you pre-saved the ROIs. Otherwise, choose the one of interest.
* After detection, filter based on "Set filters on spots" (Above and Below) to get rid of cells close to the border.
* As tracker, select "Overlap tracker", then in the next window select: "Precise", "Min IoU = 0" and "Scale factor = 1.2".
* Save the tracking as .xml file, the edges as .csv file and export ROIs to ROIManager. Name the .xml and .csv files as the chamber names.

*/

if (runtrackmate == "Yes") {
	
	Dialog.create("Before starting...");
	Dialog.addMessage("For faster and more automated analysis, it is recommended to run StarDist2D first, and then just import the ROIs.\nHowever, it is also possible to run StarDist2D inside TrackMate.\nChoose if you want to import pre-made ROIs (previous step) or generate them inside TrackMate");
	Dialog.addChoice("Use pre-made ROIs?", YesNo);
	Dialog.addMessage("It is possible to save the stacks associated to the XML file TrackMate will generate.\nThis would make easier editing the lineage in a second moment, upon need.\nHowever, saving additional stacks will use storage space.\nNote that it is very easy to add the image information to the XML file if needed (and save storage space!)");
	Dialog.addChoice("Save image?", YesNo);
	Dialog.show();
	preROIs = Dialog.getChoice();
	saveXMLimage = Dialog.getChoice();
	
	if (saveXMLimage == "No") {
		Dialog.create("How to add a stack to the XML file");
		Dialog.addMessage("Split the channels from the hyperstack previously saved.\nSave the phase contrast stack as .tif..\nOpen the TrackMate XML file with a text editor.\nLook for 'ImageData filename'.\nReplace the name with the one just saved (e.g.: 'XY65.tif').\nAdd the folder path if the file is not in the same folder as the XML file.\nSave the XML and load the TrackMate file manually to Fiji.");
		Dialog.addMessage("Remember to re-export the .csv with the tracking after editing the edges!");
		Dialog.show();
	}
	
	list_stacks = getFileList(output + "Hyperstacks/");
	
	for (i = 0; i < list_stacks.length; i++){
		print("Analysing stack " + i + 1 + " of " + list_stacks.length);
		
		//Open Hyperstack and ROIs.
		open(output + "Hyperstacks/" + list_stacks[i]);
		rename(replace(getTitle(), ".tif", ""));
		name = getTitle();
		chamber = substring(name, indexOf(name, chamberidentifier), lastIndexOf(name, ""));	
		
		if(preROIs == "Yes") roiManager("Open", output + "ROIs_StarDist2D/RoiSet_" + name + ".zip");
		
		//Split Channels.
		run("Duplicate...", "duplicate");
		rename("temporary");
		selectWindow(name);
		close();
		selectWindow("temporary");
		if (num_channels == 1) rename("C1");
		if (num_channels > 1) {
			run("Split Channels");
			for (k = 1; k <= num_channels; k++){	
				selectWindow("C" + k + "-temporary");
				rename("C" + k);
			}
		}
		
		//Turn phase-contrast to label image
		if(preROIs == "Yes") {
			selectWindow("C" + phase_channel);
			run("ROIs to Label image");
			selectImage("ROIs2Label_C" + phase_channel);
			rename("ROIs2Label");
			roiManager("reset");
			selectWindow("ROIs2Label");
		} else {
			selectWindow("C" + phase_channel);
		}
	
		//Run TrackMate.
		if (saveXMLimage == "Yes") {
			rename(date + "_" + strain + "_" + sensor + "_" + chamber);
			nameXML = getTitle();
			saveAs("Tiff", output + "XML/" + name + ".tif");
		}
	
		print("Analysing Chamber: " + chamber);
		print("Check Macro for info on which detector and tracker to use!");
		
		run("TrackMate");
		
		waitForUser("Check Macro for info on which detector and tracker to use!\nWhat should you save from TrackMate?\n-Tracking as .xml file (press on 'Save').\n-Edges as .csv file (click on 'Tracks').\n-Export ROIs to ROIManager (All spots).\n \nOnce you saved all the files, you can close TrackMate,\nBUT NOT the phase contrast image!");
		waitForUser("Done? Press 'OK' once you are done with TrackMate!");
	
		//Restore the position information in the ROIs so that all the pictures in a stack are analysed.
		for (k = 0; k < roiManager("count"); k++) {
			roiManager("Select", k);
			Roi.getPosition(channel, slice, frame);
			RoiManager.setPosition(frame);
			}
		
		//Save the ROIs	
		roiManager("Save", output + "ROIs_TrackMate/RoiSet_" + name + ".zip");
   
		//Prepare "Fresh Start" for next loop.
		roiManager("reset");
		run("Clear Results");
		print("\\Clear");
		close("*");
	}
}


/*
-------------------------------------------------
# IMAGE ANALYSIS.

Hyperstacks are analysed, so that information is saved into an excel file.
ROIs used can be chosen at the beginning of the analysis. The "Result" files will be different based on the ROIs selected (from STarDist2D or TrackMate).

Within the same "Result" file, there will be 2 sheet for each chamber: 
* One containing information about the phase contrast channel (_phase).
* One containing information about the fluorescence channels (_fluo).

In addition, there will be a "Summary_Result" file, in which there will be one sheet for each chamber. In each sheet, there will be:
* The number of cells in each fraeme.
* Mean and sd among the cells within each frame for all the fluorescent channels + fluorescence ratios. 

*/

if (runanalysis == "Yes") { 
	
	Dialog.create("Ratio Calculation");
	Dialog.addChoice("Which ROIs do you want to use?", ROIchoice);
	
	if (sensor == "No Biosensor") {
		Dialog.addMessage("No biosensor was selected at start");
		Dialog.show();
		
		chosenROIs = Dialog.getChoice();
		
		} else {
			
		Dialog.addMessage(" \nChoose fluorescence ratios to compute\n(Leave empty if not interested)");
		Dialog.addString("Ratio_Name#1", " "); Dialog.addToSameRow(); Dialog.addChoice("NUM", channelchoice); Dialog.addToSameRow(); Dialog.addNumber("Channel #", 0);
		Dialog.addToSameRow(); Dialog.addChoice("DEN", channelchoice); Dialog.addToSameRow(); Dialog.addNumber("Channel #", 0);
		
		Dialog.addString("Ratio_Name#2", " "); Dialog.addToSameRow(); Dialog.addChoice("NUM", channelchoice); Dialog.addToSameRow(); Dialog.addNumber("Channel #", 0);
		Dialog.addToSameRow(); Dialog.addChoice("DEN", channelchoice); Dialog.addToSameRow(); Dialog.addNumber("Channel #", 0);
		
		Dialog.addString("Ratio_Name#3", " "); Dialog.addToSameRow(); Dialog.addChoice("NUM", channelchoice); Dialog.addToSameRow(); Dialog.addNumber("Channel #", 0);
		Dialog.addToSameRow(); Dialog.addChoice("DEN", channelchoice); Dialog.addToSameRow(); Dialog.addNumber("Channel #", 0);
		Dialog.show();
		
		chosenROIs = Dialog.getChoice();
			
		Ratio_1 = Dialog.getString(); 
		NUM1 = Dialog.getChoice(); Nchannel1= Dialog.getNumber();
		DEN1 = Dialog.getChoice(); Dchannel1 = Dialog.getNumber();
		
		Ratio_2 = Dialog.getString(); 
		NUM2 = Dialog.getChoice(); Nchannel2= Dialog.getNumber();
		DEN2 = Dialog.getChoice(); Dchannel2 = Dialog.getNumber();
		
		Ratio_3 = Dialog.getString(); 
		NUM3 = Dialog.getChoice(); Nchannel3= Dialog.getNumber();
		DEN3 = Dialog.getChoice(); Dchannel3 = Dialog.getNumber();
		
		Ratio_Names = Array.deleteValue(newArray(Ratio_1, Ratio_2, Ratio_3), " ");
		NUMs = Array.deleteValue(newArray(NUM1, NUM2, NUM3), " ");
		Nchan = Array.deleteValue(newArray(Nchannel1, Nchannel2, Nchannel3), 0);
		DENs = Array.deleteValue(newArray(DEN1, DEN2, DEN3), " ");
		Dchan = Array.deleteValue(newArray(Dchannel1, Dchannel2, Dchannel3), 0);
		}
	
	//Saving experimental information
	print("Ratios: ");
	Array.print(Ratio_Names);
	print("ROIs from: " + chosenROIs);
	print("Numerators: ");
	Array.print(NUMs);
	print("Numerators Channels: ");
	Array.print(Nchan);
	print("Denominators: "); 
	Array.print(DENs);
	print("Denominators Channels: ");
	Array.print(Dchan);
	selectWindow("Log");
	saveAs("Text", output + "Experimental_Details#2.txt");
	
	list_stacks = getFileList(output + "Hyperstacks/");
	
	for (i = 0; i < list_stacks.length; i++){
		print("Analysing stack " + i + 1 + " of " + list_stacks.length);
		
		//Open Hyperstack.
		filename = output + "Hyperstacks/" + list_stacks[i];
		open(filename);
		rename(replace(getTitle(), ".tif", ""));
		name = getTitle();
		chamber = substring(name, indexOf(name, chamberidentifier), lastIndexOf(name, "")); 
	
		//Split Channels.
		run("Duplicate...", "duplicate");
		rename("temporary");
		selectWindow(name);
		close();
		selectWindow("temporary");
		if (num_channels > 1) run("Split Channels");
		
		for (k = 1; k <= num_channels; k++){	
			selectWindow("C" + k + "-temporary");
			rename("C" + k);
			}
	
		//Import ROIs
		roiManager("Open", output + "/ROIs_" + chosenROIs + "/RoiSet_" + name + ".zip");
		
		//Create an array to select all the ROIs and save them.
		ROIarray = newArray(roiManager("count"));
		for (k = 0; k < roiManager("count"); k += 1) { ROIarray[k] = k; }
	
		//Analyse the ROIs in the phase contrast for morphology analysis.
		run("Set Measurements...", "area mean standard perimeter shape feret's stack display redirect=C" + phase_channel + " decimal=3"); 
		roiManager("select", ROIarray);
		roiManager("Measure");
		run("Read and Write Excel", "dataset_label=[" + chamber + "] no_count_column file=[" + output + "Results/Results_" + chosenROIs + "_" + date + "_" + strain + "_" + sensor + ".xlsx] sheet=[" + chamber + "_phase]");
	 	run("Clear Results");
	 	
	 	//Analysis of cell number (if no biosensor present, otherwise implemented with fluorescence analysis).
	 	if (num_channels == 1) {
	 		nFrames = nSlices;
	 		nROIs = roiManager("count"); // Number of ROIs measured
	 		frameIndices = newArray(roiManager("count")); 
	 			 		
	 		for (q = 0; q < nROIs; q++) {
	 			frameIndex = getResult("Slice", q) - 1; // Get the slice index (0-based)
	 			frameIndices[q] = frameIndex;
	 			}
	 		
	 		counts = newArray(nFrames);
	 		
	 		//Count Number of cells in each slice
	 		for (q = 0; q < nROIs; q++) {
	 			frameIndex = frameIndices[q];
	 			counts[frameIndex]++; // Count the number of ROIs per slice
	 			}
	 			
	 		numCells = newArray(nFrames); // Array to store number of cells per slice
	 		
	 		//Number of cells for each slice
	 		for (slice = 0; slice < nFrames; slice++) {
	 			if (counts[slice] > 0) {
	 				slices[slice] = slice + 1;
	 				numCells[slice] = counts[slice]; // Store the number of cells (ROIs) for this slice
	 				} else {
	 					slices[slice] = slice + 1;
	 					numCells[slice] = 0;
	 				}}
	 				
	 		//Display the ResultsTable
	 		close("Results");
			Table.create("Results"); //Create a new Result table for the summary
				
	 		setResult("Frame", 0, "Frame");
	 		setResult("nCells", 0, "nCells"); 
	 		
	 		for (slice = 0; slice < nFrames; slice++) {
	 			setResult("Frame", slice + 1, slices[slice]);
	 			setResult("nCells", slice + 1, numCells[slice]); 
	 			}
	 		
	 		updateResults();
	 		run("Read and Write Excel", "dataset_label=[" + chamber + "] no_count_column file=[" + output + "Results/Summary_FIJIresults_" + chosenROIs + "_" + date + "_" + strain + "_" + sensor + ".xlsx] sheet=[" + chamber + "]");
	 		}	 	
	 	
	 	//Analysis of fluorescent channels.
	 	if (num_channels > 1) {
	 		
	 		//Analyse the ROIs in the fluorescent channels.
	 		for (k = 1; k <= num_channels; k++) {
	 			if (k != phase_channel) {
	 				
	 				//Prepare arrays to store fluorescence data for summary file
	 				selectWindow("C" + k);
	 				nFrames = nSlices; //Number of frames
	 				nROIs = roiManager("count"); //Number of ROIs measured
	 				roiFluo = newArray(roiManager("count")); 
	 				frameIndices = newArray(roiManager("count"));
	 				
	 				//Measure the fluorescence for all ROIs
	 				run("Set Measurements...", "mean standard stack display redirect=C" + k + " decimal=3"); 
	 				roiManager("select", ROIarray);
	 				roiManager("Measure");
	 				
	 				//Get info for summary
	 				for (x = 0; x < nROIs; x++) {
	 					frameIndex = getResult("Slice", x) - 1; // Get the slice index (0-based)
    					meanValue = getResult("Mean", x); // Get the mean fluorescence value
    					roiFluo[x] = meanValue;
    					frameIndices[x] = frameIndex;
    					}
    					
    				//Prepare arrays to store the sum of fluorescence values, the count of ROIs, and values for SD calculation for each slice
					sumFluo = newArray(nFrames);
					sumSquares = newArray(nFrames);
					counts = newArray(nFrames);
					
					//Sum fluorescence values for each slice
					for (q = 0; q < nROIs; q++) {
					    frameIndex = frameIndices[q];
					    value = roiFluo[q];
					    sumFluo[frameIndex] += value; //Sum the fluorescence values
					    sumSquares[frameIndex] += value * value; //Sum the squares of fluorescence values (for variance computations)
					    counts[frameIndex]++; //Count the number of ROIs per slice
					}

					//Prepare arrays to store results for the summary table
					slices = newArray(nFrames);
					avgFluo = newArray(nFrames);
					stdDevs = newArray(nFrames);
					numCells = newArray(nFrames); 
					
					//Calculate the average fluorescence, standard deviation (as square root of variance), and number of cells for each slice
					for (slice = 0; slice < nFrames; slice++) {
					    if (counts[slice] > 0) {
					        slices[slice] = slice + 1;
					        avgFluo[slice] = sumFluo[slice] / counts[slice];;
					        stdDevs[slice] = sqrt((sumSquares[slice] - (sumFluo[slice] * sumFluo[slice] / counts[slice])) / (counts[slice] - 1));
					        numCells[slice] = counts[slice]; 
					    } else {
					        slices[slice] = slice + 1;
					        avgFluo[slice] = NaN;
					        stdDevs[slice] = NaN;
					        numCells[slice] = 0;
					    	}
						}
					
					//Display the ResultsTable
					Table.rename("Results", "Results_tmp"); //Rename the overall result table with individual values so that one big dataframe with all the data can be saved
					Table.create("Results"); //Create a new Result table for the summary

					setResult("Frame", 0, "Frame");
					setResult("mean_C" + k, 0, "mean_C" + k);
					setResult("sd_C" + k, 0, "sd_C" + k);
					setResult("nCells", 0, "nCells"); 
					
					for (slice = 0; slice < nFrames; slice++) {
					    setResult("Frame", slice + 1, slices[slice]);
					    setResult("mean_C" + k, slice + 1, avgFluo[slice]);
					    setResult("sd_C" + k, slice + 1, stdDevs[slice]);
					    setResult("nCells", slice + 1, numCells[slice]); 
						}
					updateResults();
					
					//Save the summary table only!
					run("Read and Write Excel", "dataset_label=[C" + k + "] no_count_column file=[" + output + "Results/Summary_FIJIresults_" + chosenROIs + "_" + date + "_" + strain + "_" + sensor + ".xlsx] sheet=[" + chamber + "_fluo]");
					
					close("Results");
					Table.rename("Results_tmp", "Results");
					}}
	
			//Analyse the fluorescence ratios.
			if (Ratio_Names.length != 0) {
				
				//Create the ratio stack
				for (k = 0; k < Ratio_Names.length; k++) {
					imageCalculator("Divide create stack", "C" + Nchan[k] + "", "C" + Dchan[k] + "");
					rename(Ratio_Names[k]);
					}
				
				//Measure fluorescence in the ratio stack.
				for (k = 0; k < Ratio_Names.length; k++){
					run("Set Measurements...", "mean standard stack display redirect=" + Ratio_Names[k] + " decimal=3"); 
					roiManager("select", ROIarray);
					roiManager("Measure");
					}
				}
		
			//Fluorescence Results.
			run("Read and Write Excel", "dataset_label=[" + chamber + "] no_count_column file=[" + output + "Results/Results_" + chosenROIs + "_" + date + "_" + strain + "_" + sensor + ".xlsx] sheet=[" + chamber + "_fluo]");
			close("Results");
			
			//Summary of fluorescent ratio stacks
			for (k = 0; k < Ratio_Names.length; k++){
				selectWindow(Ratio_Names[k]);
				nFrames = nSlices; //Number of frames
	 			nROIs = roiManager("count"); //Number of ROIs measured
	 			roiFluo = newArray(roiManager("count")); 
	 			frameIndices = newArray(roiManager("count"));
				
				//Measure the fluorescence for all ROIs
				run("Set Measurements...", "mean standard stack display redirect=" + Ratio_Names[k] + " decimal=3"); 
				roiManager("select", ROIarray);
				roiManager("Measure");
				
				//Get info for summary
	 			for (x = 0; x < nROIs; x++) {
	 				frameIndex = getResult("Slice", x) - 1; // Get the slice index (0-based)
    				meanValue = getResult("Mean", x); // Get the mean fluorescence value
    				roiFluo[x] = meanValue;
    				frameIndices[x] = frameIndex;
    				}
    					
    			//Prepare arrays to store the sum of fluorescence values, the count of ROIs, and values for SD calculation for each slice
				sumFluo = newArray(nFrames);
				sumSquares = newArray(nFrames);
				counts = newArray(nFrames);
					
				//Sum fluorescence values for each slice
				for (q = 0; q < nROIs; q++) {
				    frameIndex = frameIndices[q];
				    value = roiFluo[q];
				    sumFluo[frameIndex] += value; //Sum the fluorescence values
				    sumSquares[frameIndex] += value * value; //Sum the squares of fluorescence values (for variance computations)
				    counts[frameIndex]++; //Count the number of ROIs per slice
					}

				//Prepare arrays to store results for the summary table
				slices = newArray(nFrames);
				avgFluo = newArray(nFrames);
				stdDevs = newArray(nFrames);
				numCells = newArray(nFrames); 
					
				//Calculate the average fluorescence, standard deviation (as square root of variance), and number of cells for each slice
				for (slice = 0; slice < nFrames; slice++) {
				    if (counts[slice] > 0) {
				        slices[slice] = slice + 1;
				        avgFluo[slice] = sumFluo[slice] / counts[slice];;
				        stdDevs[slice] = sqrt((sumSquares[slice] - (sumFluo[slice] * sumFluo[slice] / counts[slice])) / (counts[slice] - 1));
				        numCells[slice] = counts[slice]; 
					    } else {
					    	slices[slice] = slice + 1;
					        avgFluo[slice] = NaN;
					        stdDevs[slice] = NaN;
					        numCells[slice] = 0;
					    }}
					
				//Display the ResultsTable
				close("Results");
				Table.create("Results"); //Create a new Result table for the summary
				
				setResult("Frame", 0, "Frame");
				setResult("mean_" + Ratio_Names[k], 0, "mean_" + Ratio_Names[k]);
				setResult("sd_" + Ratio_Names[k], 0, "sd_C" + Ratio_Names[k]);
				setResult("nCells", 0, "nCells"); 
					
				for (slice = 0; slice < nFrames; slice++) {
				    setResult("Frame", slice + 1, slices[slice]);
				    setResult("mean_" + Ratio_Names[k], slice + 1, avgFluo[slice]);
				    setResult("sd_" + Ratio_Names[k], slice + 1, stdDevs[slice]);
				    setResult("nCells", slice + 1, numCells[slice]); 
					}
				updateResults();
					
				run("Read and Write Excel", "dataset_label=[" + Ratio_Names[k] + "] no_count_column file=[" + output + "Results/Summary_FIJIresults_" + chosenROIs + "_" + date + "_" + strain + "_" + sensor + ".xlsx] sheet=[" + chamber + "_fluo]");
					
				close("Results");
				close("Results_tmp");
				}

			//Save Stacks with Ratios.
	 		for (k = 0; k < Ratio_Names.length; k++){
				selectWindow(Ratio_Names[k]);
				run("Scale Bar...", "width=10 height=10 font=28 color=White background=Black location=[Upper Right] bold overlay label");
				saveAs("Tiff", output + "Hyperstacks_fluo/" + Ratio_Names[k] + "_" + name + ".tif");
				}
	 		}
	   
		//Prepare "Fresh Start" for next loop.
		roiManager("reset");
		run("Clear Results");
		print("\\Clear");
		close("*");
	}
}


beep();
print("DONE! :D");

/*
-------------------------------------------------
__Done, good job!__ You can continue the analysis in R now!

*/