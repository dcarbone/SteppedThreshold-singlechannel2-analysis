//made by Bea Carbone - 2016

importClass(Packages.java.io.File)

// reusable GenericDialog var
var gd;

// reusable DirectoryChooser var
var dc;

// Directories
var sourceDir, punctaDir, thresholdDir, resultDir;

// List of files under sourceDir
var sourceList;

// Chosen color channel
var channel;

// send up a "hello" message
gd = new GenericDialog('Instructions');
gd.addMessage("This macro needs several folders: \n-one with the separate channel images in subfolders for conditions \n-an empty one for the puncta \n-an empty one for masks \n \n -Files need to be single channel images-");
gd.showDialog();

if (gd.wasCanceled()) {
    throw new Exception('Execution cancelled');
}

// locate sourceDir
dc = new DirectoryChooser('Choose folder for images');
sourceDir = dc.getDirectory();

// get list of files under this directory
var dir = new java.io.File(sourceDir);
sourceList = dir.listFiles();

// determine which color to look for
gd = new GenericDialog('Channel Color');
gd.addChoice('Choose one:', ['Green', 'Red', 'Blue', 'Cyan'], 'Green');
gd.showDialog();
channel = gd.getNextChoice();

// get puncta output dir
dc = new DirectoryChooser('Choose folder for ' + channel + ' puncta');
punctaDir = dc.getDirectory();

// get threshold output dir
dc = new DirectoryChooser('Choose folder for ' + channel + ' masks');
thresholdDir = dc.getDirectory();

// get excel output dir
dc = new DirectoryChooser('Choose folder for Excel files');
resultDir = dc.getDirectory();

// Do this...
IJ.run("Set Measurements...", "area mean integrated area_fraction limit redirect=None decimal=3");

var l = "";
for (var t = 0; t < 160; t += 3) {
    l = l + "	" + t;
}

ROIstring = "puncta/um^2	average_area	average_intensity\nthresholding	area_ofROI(um^2)" + l + "		thresholding" + l + "		thresholding" + l + "\n";

allAreaArray = [];
allPunctaArray = [];

nameArray = [];

var statsPerROIarea = ["dendrite", "min", "max", "mean", "stdDev"];
var statsperROIpuncta = ["dendrite", "min", "max", "mean", "stdDev"];

threshStatsPuncta = ["dendrite", "threshold", "min", "max", "mean", "stdDev"];
threshStatsArea = ["dendrite", "threshold", "min", "max", "mean", "stdDev"];

PunctaStandardDevArray = [];
PunctaStandardMeanArray = [];
PunctaStandardMaxArray = [];

AreastandardDevArray = [];
AreastandardMeanArray = [];
AreastandardMaxArray = [];

condition = getFileList(sourceDir);
Dialog.create("which condition should be used for analysis standardization?");
Dialog.addChoice("Condition", condition);
Dialog.show();
standard = Dialog.getChoice();

var areas = [], punctaThresholds = [];
var corNum;

e = 0;
for (i = 0; i < sourcelist.length; i++) {
    areaROIarray = [];
    condition = getFileList(sourceDir + sourcelist[i]);
    filename = File.getName(sourceDir + sourcelist[i]);
    allAreaArray = Array.concat(allAreaArray, filename);
    allPunctaArray = Array.concat(allPunctaArray, filename);
    statsPerROIarea = Array.concat(statsPerROIarea, filename);
    statsperROIpuncta = Array.concat(statsperROIpuncta, filename);
    threshStatsPuncta = Array.concat(threshStatsPuncta, filename);
    threshStatsArea = Array.concat(threshStatsArea, filename);

    punctastats = [];
    areastats = [];

    ROIstring = ROIstring + filename + "\n";
    Dialog.create("list the area per ROI");
    for (p = 0; p < condition.length; p++) {
        Dialog.addNumber(condition[p], 1);
    }
    Dialog.show;
    for (p = 0; p < condition.length; p++) {
        areaROI = Dialog.getNumber();
        areaROIarray = Array.concat(areaROIarray, areaROI);
    }

    for (d = 0; d < condition.length; d++) {
        tempAreaArray = [];
        tempPunctaArray = [];
        CorrectedPunctaStr = CorrectedPunctaStr + filename;
        open(sourceDir + sourcelist[i] + condition[d]);
        title = getTitle();
        run("Properties...", "channels=1 slices=1 frames=1 unit=um pixel_width=0.1705 pixel_height=0.1705 voxel_depth=1");
        run("Gaussian Blur...", "sigma=1");
        setColor("Magenta");
        _int = title;
        Num = title;
        are = title;
        allAreaArray.push("");
        allPunctaArray.push("");
        corNum = title;
        corNum += "	" + areaROIarray[d];
        for (t = 0; t < 160; t += 3) {
            selectImage(title);
            shortTitle = substring(title, 0, lengthOf(title) - 4);
            nameArray.push(shortTitle + "_" + t);
            namestring += shortTitle + "_" + t;
            mask();
            Num = Num + "	" + number;
            are = are + "	" + area;
            _int = _int + "	" + intensity;
            corNum = corNum + "	" + (number / areaROIarray[d]);

            if (t % 3 === 0) {
                if (typeof areas[t] === 'undefined') {
                    areas[t] = [];
                }

                if (typeof punctaThresholds[t] === 'undefined') {
                    punctaThresholds[t] = [];
                }

                areas[t].push(number);
                punctaThresholds.push(area);
            }
        }
        ROIstring = ROIstring + corNum + "		" + are + "		" + _int + "\n";
        CorrectedPunctaStr = CorrectedPunctaStr + corNum;
        Array.getStatistics(tempAreaArray, min, max, mean, stdDev);
        statsPerROIarea = Array.concat(statsPerROIarea, nameArray[e] + "," + min + "," + max + "," + mean + "," + stdDev);
        Array.getStatistics(tempPunctaArray, min, max, mean, stdDev);
        statsperROIpuncta = Array.concat(statsperROIpuncta, nameArray[e] + "," + min + "," + max + "," + mean + "," + stdDev);

        selectImage(title);
        close();
        selectWindow(title + "_" + channel + 0 + "_mask");
        run("Select None");
        setSlice(1);
        run("Label...", "format=Label starting=0 interval=1 x=5 y=5 font=8 text=[] range=1-51 use");
        saveAs("Tiff", thresholdDir + shortTitle);
        close();
        e = e + 54;
    }

    for (var ii = 0; ii <= 160; ii += 3) {
        Array.getStatistics(punctaThresholds[ii], min, max, mean, stdDev);
        threshStatsPuncta.push(nameArray[e - 54] + ',' + min + "," + max + "," + mean + "," + stdDev)
        if (sourcelist[i] === standard) {
            PunctaStandardMaxArray.push(max);
            PunctaStandardMeanArray.push(mean);
            PunctaStandardDevArray.push(stdDev);
        }
    }
}

Array.getStatistics(AreastandardMeanArray, min, max, mean, stdDev);
StatsArea = max;
StatsDevArea = stdDev;
Array.getStatistics(PunctaStandardMeanArray, min, max, mean, stdDev);
var StatsPuncta = max;
StatsDevPuncta = stdDev;

for (v = 0; v < AreastandardMeanArray.length; v++) {
    if (AreastandardMeanArray[v] === StatsArea) {
        stopArea = v;
    }
    else if (PunctaStandardMeanArray[v] == StatsPuncta) {
        stopPuncta = v;
    }
}
if (stopPuncta > stopArea) {
    AnalysisPuncta = Array.slice(PunctaStandardMeanArray, stopPuncta - stopArea, stopPuncta);
    AnalysisArea = Array.slice(AreastandardMeanArray, stopPuncta - stopArea, stopPuncta);
    //the threshold should be ~1/2 between the difference of the two - multiplied by 3 to get threshold equivalent
    var analyze = (stopPuncta - (round((stopPuncta - stopArea) / 4))) * 3;
    waitForUser("Use data from threshold " + analyze);
    ROIstring = ROIstring + "use threshold " + analyze;
}


File.saveString(ROIstring, resultDir + channel + ".xls");

function mask() {
    tempTitle = getTitle();
    run("Clear Results");
    run("Threshold...");
    setAutoThreshold("Default dark");
    setThreshold(t, 255);
    selectWindow(tempTitle);
    run("Find Maxima...", "noise=7.5 output=[Segmented Particles] above");
    selectImage(tempTitle + " Segmented");
    run("Watershed");
    run("Analyze Particles...", "size=1-60 pixel show=Masks add");
    number = roiManager("count");
    roiManager("Measure");
    area = 0;
    if (roiManager("count") > 0) {
        roiManager("Save", punctaDir + "ROI_" + shortTitle + "_" + t + ".zip");
        for (j = 0; j < nResults; j++) {
            area = area + getResult("Area", j);
            intensity = intensity + getResult("IntDen", j);
        }
        area = area / nResults;
        intensity = intensity / nResults;
    }
    roiManager("reset");
    selectImage("Mask of " + tempTitle + " Segmented");
    rename(title + "_" + channel + t + "_mask");
    selectImage(tempTitle + " Segmented");
    close();
    if (t > 0) {
        selectImage(title + "_" + channel + t + "_mask");
        run("Select All");
        run("Copy");
        close();
        selectImage(title + "_" + channel + 0 + "_mask");
        run("Add Slice");
        setSlice(nSlices);
        shortTitle = substring(title, 0, lengthOf(title) - 4);
        run("Set Label...", "label=" + shortTitle + "_" + t);
        run("Paste");
    }
}
