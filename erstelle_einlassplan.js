/**
 * Testscript
 * @author Adrian Kauz
 */

var sScriptName = "FUEP-Script"
var oArgs = WScript.Arguments;
var oCom;

main();

 /*
================
main()
================
*/
function main()
{
	oCom = new ComObjects();

	if(!oCom.loadAllObjects()) {
		WScript.echo("Konnte nicht alle COM-Objekte laden!");
		return;
	}

	if(oArgs.length !== 1) {
		showExclamationBox("No Arguments!")
		return;
	}

	if(isCSVFile(oArgs(0)) === false) {
		showExclamationBox("Keine oder ungueltige Datei!")
		return;
	}

	var arrAllMovies = loadAllMovies(oArgs(0));

	for(var x = 0; x < arrAllMovies.length; x++){
        showInfoBox(arrAllMovies[x].toString());
	}


	return;
	oCom.Excel.Visible = true;
	oCom.Excel.Workbooks.Open(oArgs(0));
}


/*
================
ComObjects()
================
*/
function ComObjects()
{
	this.Shell = null;
	this.FSO = null;
	this.Excel = null;

	this.loadAllObjects = function() {
		try{
			this.Shell = new ActiveXObject("WScript.Shell");
			this.FSO = new ActiveXObject("Scripting.FileSystemObject");
			this.Excel = new ActiveXObject("Excel.Application");
			return true;
		} catch(ex) {
			return false;
		}
	}
}

 /*
================
showExclamationBox()
================
*/
function showExclamationBox(sMessage)
{
	oCom.Shell.popup(sMessage, 0, sScriptName, 48 );
}

/*
================
showInfoBox()
================
*/
function showInfoBox(sMessage)
{
    oCom.Shell.popup(sMessage, 0, sScriptName, 64 );
}

 /*
================
isCSVFile()
================
*/
function isCSVFile(sPath)
{
	if(oCom.FSO.fileExists(oArgs(0))){
		if(sPath.toLowerCase().indexOf(".csv") !== -1) {
			return true;
		}
	}

	return false;
}

 /*
================
parseCSV()
================
*/
function loadAllMovies(sPath)
{
    var regex = /(0[1-9]|[12][0-9]|3[01])[.](0[1-9]|1[012])[.](20)\d\d/g;
	var oFile = oCom.FSO.OpenTextFile(sPath, 1);
	var arrMovieObjects = [];

	while(!oFile.AtEndOfStream) {
		var arrCurrLine = oFile.ReadLine().split(",");

		if(arrCurrLine[0].match(regex) !== null){
			var movieObject = new MovieObject();

			movieObject.Datum				= arrCurrLine[0];
            movieObject.SaalName			= arrCurrLine[1];
            movieObject.SaalNummer			= arrCurrLine[2];
            movieObject.Schiene				= arrCurrLine[3];
            movieObject.VorstellungStart	= arrCurrLine[4];
            movieObject.HauptfilmStart		= arrCurrLine[5];
            movieObject.Pause				= arrCurrLine[6];
            movieObject.PauseEnde			= arrCurrLine[7];
            movieObject.Dialicht			= arrCurrLine[8];
            movieObject.Ende				= arrCurrLine[9];
            movieObject.Format				= arrCurrLine[10];
            movieObject.Filmtitel			= arrCurrLine[11];
            movieObject.FSK					= arrCurrLine[12];
            movieObject.Aufraumzeit			= arrCurrLine[13];
            movieObject.NV					= arrCurrLine[14];
            movieObject.NVFormat			= arrCurrLine[15];
            movieObject.NVFilmTitel			= arrCurrLine[16];
            arrMovieObjects.push(movieObject);
		}
	}

	return arrMovieObjects;
}

/*
================
MovieObject()
================
*/
function MovieObject()
{
    this.Datum = null;
    this.SaalName = null;
    this.SaalNummer = null;
    this.Schiene = null;
    this.VorstellungStart = null;
    this.HauptfilmStart = null;
    this.Pause = null;
    this.PauseEnde = null;
    this.Dialicht = null;
    this.Ende = null;
    this.Format = null;
    this.Filmtitel = null;
    this.FSK = null;
    this.Aufraumzeit = null;
    this.NV = null;
    this.NVFormat = null;
    this.NVFilmTitel = null;

    this.toString = function(){
    	return "Datum:\t\t" + this.Datum
			+ "\nSaalName:\t" + this.SaalName
            + "\nSaalNummer:\t" + this.SaalNummer
            + "\nSchiene:\t\t" + this.Schiene
            + "\nVorstellungStart:\t" + this.VorstellungStart
            + "\nHauptfilmStart:\t" + this.HauptfilmStart
            + "\nPause:\t\t" + this.Pause
            + "\nPauseEnde:\t" + this.PauseEnde
            + "\nDialicht:\t\t" + this.Dialicht
            + "\nEnde:\t\t" + this.Ende
            + "\nFormat:\t\t" + this.Format
            + "\nFilmtitel:\t\t" + this.Filmtitel
            + "\nFSK:\t\t" + this.FSK
            + "\nAufräumzeit:\t" + this.Aufraumzeit
            + "\nNV:\t\t" + this.NV
            + "\nNVFormat:\t" + this.NVFormat
            + "\nNVFilmTitel:\t" + this.NVFilmTitel;
	}
}