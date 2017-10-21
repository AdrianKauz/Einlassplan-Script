/**
 * Einlassplan-Script
 * @author Adrian Kauz
 */

var strScriptFullPath = WScript.ScriptFullName;
var strScriptPath = strScriptFullPath.substring(0, strScriptFullPath.lastIndexOf("\\"));
var strScriptName = "Einlassplan-Script";
var strVorlage = "einlassplan_vorlage.xlsx";
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
        showExclamationBox("No Arguments!");
        return;
    }

    if(isCSVFile(oArgs(0)) === false) {
        showExclamationBox("Keine oder ungueltige Datei!");
        return;
    }

    //var arrScreenings = loadScreeningsFromCSV(oArgs(0));




/*
    for(var x = 0; x < arrAllMovies.length; x++){
        showInfoBox(arrAllMovies[x].toString());
    }
*/
    oCom.Excel.Visible = true;
    //oWorkbook = oCom.Excel.Workbooks.Open(strScriptPath + "\\" + strVorlage);
	oWorkbook = oCom.Excel.Workbooks.Open(oArgs(0));

    oWorkSheet = oWorkbook.ActiveSheet;

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
    };
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
loadScreeningsFromCSV()
================
*/
function loadScreeningsFromCSV(sPath)
{
    var regex = /(0[1-9]|[12][0-9]|3[01])[.](0[1-9]|1[012])[.](20)\d\d/g;
    var oFile = oCom.FSO.OpenTextFile(sPath, 1);
    var arrMovieObjects = [];

    while(!oFile.AtEndOfStream) {
        var arrCurrLine = oFile.ReadLine().split(",");

        if(arrCurrLine[0].match(regex) !== null){
            var movieObject = new MovieObject();

            movieObject.Datum                   = arrCurrLine[0];   // "Datum"
            movieObject.SaalName                = arrCurrLine[1];   // "Saal"
            movieObject.SaalNummer              = arrCurrLine[2];   // "Saal"
            movieObject.Schiene                 = arrCurrLine[3];   // "Schiene"
            movieObject.VorstellungStart        = arrCurrLine[4];   // "Vorstellung Start"
            movieObject.HauptfilmStart          = arrCurrLine[5];   // "Hauptfilm Start"
            movieObject.Pause                   = arrCurrLine[6];   // "Pause"
            movieObject.PauseEnde               = arrCurrLine[7];   // "Pause-Ende"
            movieObject.Dialicht                = arrCurrLine[8];   // "Dialicht"
            movieObject.Ende                    = arrCurrLine[9];   // "Ende"
            movieObject.Format                  = arrCurrLine[10];  // "2D/3D"
            movieObject.Filmtitel               = arrCurrLine[11];  // "Filmtitel"
            movieObject.FSK                     = arrCurrLine[12];  // "FSK"
            movieObject.Aufraumzeit             = arrCurrLine[13];  // "Aufräumzeit"
            movieObject.NV                      = arrCurrLine[14];  // "Nächste Vorstellung"
            movieObject.NVFormat                = arrCurrLine[15];  // "2D/3D"
            movieObject.NVFilmTitel             = arrCurrLine[16];  // "Filmtitel"
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

/*
================
Dictionary()
================
*/
function Dictionary()
{
    this.arrDictionary = null;

    if (this.arrDictionary === null) {
        this.arrDictionary = [];
    }

	/**
     * Add new element to the dictionary.
	 * @param {String} sKey
     * @param {Object} oObject
     * @returns {boolean}
	 */
    this.Add = function(sKey, oObject)
    {
        if (sKey !== "" && sKey.constructor === String) {
	        for (var x = 0; x < this.arrDictionary.length; x++) {
		        if (this.arrDictionary[x]._sKey === sKey) {
			        this.arrDictionary[x]._oObject = oObject;
			        return true;
		        }
	        }

	        this.arrDictionary.push({_sKey : sKey, _oObject : oObject});
	        return true;
        }

        return false;
    };


	/**
	 * If exists, return object from dictionary element.
	 * @param {String} sKey
     * @returns {Object}
	 */
	this.Get = function(sKey)
	{
		if (sKey.constructor === String) {
			for (var x = 0; x < this.arrDictionary.length; x++) {
				if(this.arrDictionary[x]._sKey === sKey) {
					return this.arrDictionary[x]._oObject;
                }
            }
        }

		return null;
	};


	/**
	 * Checks if dictionary entry already exists
	 * @param {String} sKey
     * @returns {object}
	 */
	this.ContainsKey = function(sKey)
    {
	    if (sKey.constructor === String) {
		    for (var x = 0; x < this.arrDictionary.length; x++) {
			    if(this.arrDictionary[x]._sKey === sKey) {
				    return true;
			    }
		    }

		    return false;
	    }

	    return null;
    };


	/**
	 * Returns size of dictionary
	 * @returns {number}
	 */
	this.Count = function()
    {
        return this.arrDictionary.length;
    };
}