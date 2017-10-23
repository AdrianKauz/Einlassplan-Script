/**
 * Einlassplan-Script
 * @author Adrian Kauz
 */

var sScriptFullPath = WScript.ScriptFullName;
var sScriptPath = sScriptFullPath.substring(0, sScriptFullPath.lastIndexOf("\\"));
var sScriptName = "Einlassplan-Script";
var sVorlage = "einlassplan_vorlage.xlsx";
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

    var ScreeningCollection = loadScreeningsFromCSV(oArgs(0));


    //var arrScreenings = loadScreeningsFromCSV(oArgs(0));




/*
    for(var x = 0; x < arrAllMovies.length; x++){
        showInfoBox(arrAllMovies[x].ToString());
    }

    oCom.Excel.Visible = true;
    //oWorkbook = oCom.Excel.Workbooks.Open(strScriptPath + "\\" + strVorlage);
    oWorkbook = oCom.Excel.Workbooks.Open(oArgs(0));

    oWorkSheet = oWorkbook.ActiveSheet;
*/
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
function loadScreeningsFromCSV(sPath) {
	var regex = /(0[1-9]|[12][0-9]|3[01])[.](0[1-9]|1[012])[.](20)\d\d/g;
	var oFile = oCom.FSO.OpenTextFile(sPath, 1);
	var arrColumnTitles = [];
	var iCurrLineNr = 0;
	var sCurrLine;

	while (!oFile.AtEndOfStream) {
		iCurrLineNr++;
		sCurrLine = oFile.ReadLine();

		if (sCurrLine.length > 0) {
			if (iCurrLineNr === 1) {
				arrColumnTitles = sCurrLine.split(",");
			} else {
				var arrCurrValues = sCurrLine.split(",");

				if (arrColumnTitles.length === arrCurrValues.length) {
					var oCurrScreening = new ScreeningObject();
					oCurrScreening.addValues(arrColumnTitles, arrCurrValues);
					oCurrScreening.setSortColumn("Dialicht");
					showInfoBox(oCurrScreening.toString());
					return;
				}
			}
		}
	}
}



/*
================
ScreeningObject()
================
*/
function ScreeningObject()
{
    this.dictValues = new Dictionary();
	this.sSortColumn = "Vorstellung Start";

	/**
	 * Fills up dictionary with screening attributes.
	 * @param {Array} arrNames
	 * @param {Array} arrValues
	 */
    this.addValues = function(arrNames, arrValues)
    {
        if ((arrNames !== null) && (arrValues !== null)) {
            for (var x = 0; x < arrNames.length; x++) {
                this.dictValues.add(arrNames[x], arrValues[x]);
            }
        }
    };


	/**
	 * @param {String} sNewSortColumn
	 */
    this.setSortColumn = function(sNewSortColumn)
    {
	    if (sNewSortColumn !== "" && sNewSortColumn.constructor === String) {
	        if (this.dictValues.count() > 0 ) {
	            if (this.dictValues.containsKey(sNewSortColumn)) {
	                this.sSortColumn = sNewSortColumn;
                }
            }
	    }
    };


	/**
	 * For value comparison.
	 * @param {Object} oObject1
	 * @param {Object} oObject2
	 * @returns {Number}
	 */
    this.compare = function( oObject1, oObject2 )
    {
        return 1;
    };


	/**
	 * Returns content of the screening object as String.
	 * @returns {String}
	 */
	this.toString = function()
	{
	    sNewString = "";

	    for (var x = 0; x < this.dictValues.count(); x++) {
	        sCurrKey = this.dictValues.getKey(x);
	        sNewString += "\"" + sCurrKey + "\" --> " + this.dictValues.get(sCurrKey) + "\n";
        }

		return sNewString;
	};
}



/*
================
Dictionary()
================
*/
function Dictionary()
{
    this.arrDictionary = [];


    /**
     * Add new element to the dictionary.
     * @param {String} sKey
     * @param {Object} oObject
     * @returns {boolean}
     */
    this.add = function(sKey, oObject)
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
    this.get = function(sKey)
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
    this.containsKey = function(sKey)
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
     * @param {Number} iPosition
	 * @returns {String}
	 */
    this.getKey = function(iPosition)
    {
        if (!isNaN(iPosition)) {
            if (iPosition < this.arrDictionary.length) {
                return this.arrDictionary[iPosition]._sKey;
            }
        }

        return null;
    };


    /**
     * Returns size of dictionary
     * @returns {number}
     */
    this.count = function()
    {
        return this.arrDictionary.length;
    };
}