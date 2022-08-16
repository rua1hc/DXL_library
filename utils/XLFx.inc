//<  this is necessary for DXL encryption to work. 
/*
* This file contains many common useful functions to manipulate Excel through
* DOORS DXL.  Many of the functions or ideas came from Michael Sutherlands 
* Enhanced Export to Excel program developed for Galactic Soluctions.  So 
* just giving credit where it is due.
*/

//#include <ole.inc>
//<
// Utilities for OLE interfaces

// $Header: /products/DOORS/lib/dxl/utils/ole.inc 3     5/13/01 1:33p Martin $

// Copyright © 2001 Telelogic AB. All rights reserved.
// See DOORS manuals for copying conditions.
// Copy this file to a different location before modifying it.

/*
    Utilities (constants and functions) for OLE interfaces
    Modified by David Pechacek to add functionality.
*/

//------------------------ Objects ------------------------------

const string cObjWordApplication    = "Word.Application"
const string cObjPptApplication     = "PowerPoint.Application"
const string cObjOutlookApplication = "Outlook.Application"
const string cObjExcelApplication   = "Excel.Application"

//------------------------ VBA Properties -----------------------


const string cPropertyActivePresentation        = "ActivePresentation"
const string cPropertyActiveWorkbook            = "ActiveWorkbook"
const string cPropertyActiveSheet               = "ActiveSheet"
const string cPropertyActiveWindow              = "ActiveWindow"
const string cPropertyCheckSpelling             = "CheckSpellingAsYouType"
const string cPropertyCheckGrammar              = "CheckGrammarAsYouType"
const string cPropertyEnableAutoComplete        = "EnableAutoComplete"
const string cPropertyAutoRecover               = "AutoRecover"
const string cPropertyEnabled                   = "Enabled"
const string cPropertyErrorCheckingOptions      = "ErrorCheckingOptions"
const string cPropertyBackgroundChecking        = "BackgroundChecking"
const string cPropertyAutoCorrect               = "AutoCorrect"
const string cPropertyTwoInitialCapitals        = "TwoInitialCapitals"
const string cPropertyCorrectSentenceCap        = "CorrectSentenceCap"
const string cPropertyCapitalizeNamesofDays     = "CapitalizeNamesOfDays"
const string cPropertyCorrectCapsLock           = "CorrectCapsLock"
const string cPropertyReplaceText               = "ReplaceText"
const string cPropertyDisplayAutoCorrectOptions = "DisplayAutoCorrectOptions"
const string cPropertyHorizontalAlignment       = "HorizontalAlignment"
const string cPropertyVerticalAlignment         = "VerticalAlignment"

const string cPropertySlides            = "Slides"
const string cPropertyShapes            = "Shapes"
const string cPropertyPlaceholders      = "Placeholders"
const string cPropertyWorkbooks         = "Workbooks"
const string cPropertyTitle             = "Title"
const string cPropertyTextFrame         = "TextFrame"
const string cPropertyTextRange         = "TextRange"
const string cPropertyText              = "Text"
const string cPropertyViewType          = "ViewType"
const string cPropertyView              = "View"
const string cPropertyValue             = "Value"
const string cPropertyLeft              = "Left"
const string cPropertyTop               = "Top"
const string cPropertyHeight            = "Height"
const string cPropertyWidth             = "Width"
const string cPropertyVisible           = "Visible"
const string cPropertyBody              = "Body"
const string cPropertySubject           = "Subject"
const string cPropertyAttachments       = "Attachments"
const string cPropertyCells             = "Cells"
const string cPropertyActiveDocument    = "ActiveDocument"
const string cPropertyPageSetup         = "PageSetup"
const string cPropertyPageWidth         = "PageWidth"
const string cPropertyPrintArea         = "PrintArea"
const string cPropertyLeftMargin        = "LeftMargin"
const string cPropertyApplication       = "Application"
const string cPropertyDocuments         = "Documents"
const string cPropertySelection         = "Selection"
const string cPropertyExtendMode        = "ExtendMode"
const string cPropertyCount             = "Count"
const string cPropertyOptions           = "Options"
const string cPropertyPagination        = "Pagination"
const string cPropertyParagraphs        = "Paragraphs"
const string cPropertyType              = "Type"
const string cPropertyStyle             = "Style"
const string cPropertyStyles            = "Styles"
const string cPropertyInlineShapes      = "InlineShapes"
const string cPropertyRange             = "Range"
const string cPropertyOutlineLevel      = "OutlineLevel"
const string cPropertyTables            = "Tables"
const string cPropertyHeadingFormat     = "HeadingFormat"
const string cPropertyRows              = "Rows"
const string cPropertyRow               = "Row"
const string cPropertyRowHeight         = "RowHeight"
const string cPropertyScreenUpdating    = "ScreenUpdating"
const string cPropertyColumns           = "Columns"
const string cPropertyColumn            = "Column"
const string cPropertyColumnWidth       = "ColumnWidth"
const string cPropertyBorders           = "Borders"
const string cPropertyEnable            = "Enable"
const string cPropertyBorder            = "Border"
const string cPropertyLineStyle         = "LineStyle"
const string cPropertyNext              = "Next"
const string cPropertyNameLocal         = "NameLocal"
const string cPropertyInUse             = "InUse"
const string cPropertyVersion           = "Version"
const string cPropertyListTemplate      = "ListTemplate"
const string cPropertyParagraphFormat   = "ParagraphFormat"
const string cPropertyLeftIndent        = "LeftIndent"
const string cPropertyFont              = "Font"
const string cPropertyName              = "Name"
const string cPropertySheets            = "Sheets"
const string cPropertySymbol            = "Symbol"
const string cPropertyFontStyle         = "FontStyle"
const string cPropertyRegular           = "Regular"
const string cPropertyBold              = "Bold"
const string cPropertyBoldItalic        = "Bold Italic"
const string cPropertyItalic            = "Italic"
const string cPropertyUnderline         = "Underline"
const string cPropertyStrikeThrough     = "StrikeThrough"
const string cPropertySubscript         = "Subscript"
const string cPropertySuperscript       = "Superscript"
const string cPropertyCharacters        = "Characters"
const string cPropertyInterior          = "Interior"
const string cPropertyColor             = "Color"
const string cPropertyColorIndex        = "ColorIndex"
const string cPropertyMergeCells        = "MergeCells"
const string cPropertyWrapText          = "WrapText"
const string cPropertyFormula           = "Formula"
const string cPropertySections          = "Sections"
const string cPropertySection           = "Section"
const string cPropertyNumberFormat      = "NumberFormat"
const string cPropertyUpdateLinks       = "UpdateLinks"

//---------------------- VBA Parameters --------------------------
const string cParamStart    = "Start"
const string cParamLength   = "Length"

//------------------------ VBA Methods ---------------------------

const string cMethodActivate             = "Activate"
const string cMethodAdd                  = "Add"
const string cMethodAutoFormat           = "AutoFormat"
const string cMethodPaste                = "Paste"
const string cMethodConvertToTable       = "ConvertToTable"
const string cMethodItem                 = "Item"
const string cMethodCells                = "Cells"
const string cMethodSelect               = "Select"
const string cMethodDelete               = "Delete"
const string cMethodAddPicture           = "AddPicture"
const string cMethodCreateItem           = "CreateItem"
const string cMethodDisplay              = "Display"
const string cMethodSave                 = "Save"
const string cMethodSaveAs               = "SaveAs"
const string cMethodRange                = "Range"
const string cMethodInsert               = "Insert"
const string cMethodInsertAfter          = "InsertAfter"
const string cMethodInsertBefore         = "InsertBefore"
const string cMethodInsertParagraphAfter = "InsertParagraphAfter"
const string cMethodInsertBreak          = "InsertBreak"
const string cMethodEscapeKey            = "EscapeKey"
const string cMethodCollapse             = "Collapse"
const string cMethodPasteSpecial         = "PasteSpecial"
const string cMethodEndKey               = "EndKey"
const string cMethodNew                  = "New"
const string cMethodCell                 = "Cell"
const string cMethodMoveRight            = "MoveRight"
const string cMethodNext                 = "Next"
const string cMethodTypeBackspace        = "Typebackspace"   
const string cMethodOpen                 = "Open"
const string cMethodClose                = "Close"
const string cMethodCopy                 = "Copy"
const string cMethodRepaginate           = "Repaginate"
const string cMethodRun                  = "Run"

//--------------------Microsoft Excel Color Values------------------
const int cColorAutomatic        = 0
const int cColorBlack            = 1
const int cColorBrown            = 53
const int cColorOliveGreen       = 52
const int cColorDarkGreen        = 51
const int cColorDarkTeal         = 49
const int cColorDarkBlue         = 11
const int cColorIndigo           = 55
const int cColorGrey80           = 56
const int cColorDarkRed          = 9
const int cColorOrange           = 46
const int cColorDarkYellow       = 12
const int cColorGreen            = 10
const int cColorTeal             = 14
const int cColorBlue             = 5
const int cColorBlueGray         = 47
const int cColorGray50           = 16
const int cColorRed              = 3
const int cColorLightOrange      = 45
const int cColorLime             = 43
const int cColorSeaGreen         = 50
const int cColorAqua             = 42
const int cColorLightBlue        = 41
const int cColorViolet           = 13
const int cColorGray40           = 48
const int cColorPink             = 7
const int cColorGold             = 44
const int cColorYellow           = 6
const int cColorBrightGreen      = 4
const int cColorTurquoise        = 8
const int cColorSkyBlue          = 33
const int cColorPlum             = 54
const int cColorGray25           = 15
const int cColorRose             = 38
const int cColorTan              = 40
const int cColorLightYellow      = 36
const int cColorLightGreen       = 35
const int cColorLightTurquoise   = 34   
const int cColorPaleBlue         = 37
const int cColorLavender         = 39
const int cColorWhite            = 2
const int cColorxlNone           = -4142

//----------------- Microsoft Excel Number Formats -----------------
const string cNumberFormat_General      = "General"
const string cNumberFormat_Number       = "0.00"
const string cNumberFormat_WholeNumber  = "0"
const string cNumberFormat_NumberWithSeparator = "#,##0.00"
const string cNumberFormat_Currency     = "$#,##0.00"
const string cNumberFormat_Accounting   = "_($* #,##0.00_);_($* (#,##0.00);_($* \"\"-\"\"??_);_(@_)"
const string cNumberFormat_Date1        = "m/d"               // outputs as 3/14
const string cNumberFormat_Date2        = "m/d/yy"            // outputs as 3/14/98
const string cNumberFormat_Date3        = "mm/dd/yy"          // outputs as 03/14/98
const string cNumberFormat_Date4        = "d-mmm"             // outputs as 14-Mar
const string cNumberFormat_Date5        = "d-mmm-yy"          // outputs as 14-Mar-98
const string cNumberFormat_Date6        = "mmm-yy"            // outputs as Mar-98
const string cNumberFormat_Date7        = "mmmm-yy"           // outputs as March-98
const string cNumberFormat_Date8        = "mmmm d, yyyy"      // outputs as March 14, 1998
const string cNumberFormat_Date9        = "m/d/yy h:mm AM/PM" // outputs as 3/14/98 1:30 PM
const string cNumberFormat_Date10       = "m/d/yy h:mm"      // outputs as 3/14/98 13:30
const string cNumberFormat_Date11       = "m/d/yyyy"         // outputs as 3/14/1998
const string cNumberFormat_Date12       = "d-mmm-yyyy"       // outputs as 14-Mar-1998
const string cNumberFormat_Time1        = "h:mm"              // outputs as 13:30
const string cNumberFormat_Time2        = "h:mm AM/PM"        // outputs as 1:30 PM
const string cNumberFormat_Time3        = "h:mm:ss"           // outputs as 13:30:55
const string cNumberFormat_Time4        = "h:mm:ss AM/PM"     // outputs as 1:30:55 PM
const string cNumberFormat_Percentage   = "0.00%"
const string cNumberFormat_Fraction1    = "# ?/?"         // Up to two digits (1/4)
const string cNumberFormat_Fraction2    = "# ??/??"       // Up to two digits (21/25)
const string cNumberFormat_Fraction3    = "# ???/???"     // Up to three digits (312/943)
const string cNumberFormat_Fraction4    = "# ?/2"         // As halves (1/2)
const string cNumberFormat_Fraction5    = "# ?/4"         // As quarters (2/4)
const string cNumberFormat_Fraction6    = "# ?/8"         // As eighths (4/8)
const string cNumberFormat_Fraction7    = "# ??/16"       // As sixteenths (8/16)
const string cNumberFormat_Fraction8    = "# ?/10"        // As tenths (3/10)
const string cNumberFormat_Fraction9    = "# ??/100"      // As hundreths (30/100)
const string cNumberFormat_Scientific   = "0.00E+00"
const string cNumberFormat_Text         = "@"
const string cNumberFormat_ZipCode      = "00000"
const string cNumberFormat_ZipCode4     = "00000-0000"
const string cNumberFormat_PhoneNumber  = "[<=9999999]###-####;(###) ###-####"
const string cNumberFormat_SSN          = "000-00-0000"

//----------------- Microsoft Excel Alignment Constants -----------------
const int xlHAlignCenter         = -4108
const int xlHAlignJustify        = -4130
const int xlHAlignLeft           = -4131
const int xlHAlignRight          = -4152
const int xlVAlignTop            = -4160
const int xlVAlignBottom         = -4107
const int xlVAlignCenter         = -4108
const int xlVAlignJustify        = -4130

//------------------ PowerPoint Symbolic Constants -----------------

// you can find out what these are by creating a macro in powerpoint
// including the constant and stepping into the macro
//  hold the mouse pointer over the constant and a tooltip will tell you the value

const int ppLayoutTitle             = 1
const int ppLayoutText              = 2
const int ppLayoutTwoColumnText     = 3
const int ppLayoutTable             = 4
const int ppLayoutTitleOnly         = 11
const int ppLayoutBlank             = 12
const int ppLayoutTextAndObject     = 13
const int ppLayoutObjectAndText     = 14  
const int ppLayoutLargeObject       = 15
const int ppLayoutObject            = 16      
const int ppLayoutTextAndTwoObjects = 21   
const int ppLayoutFourObjects       = 24
const int ppViewSlide               = 1
const int ppPlaceholderObject       = 7

//------------------ Outlook Symbolic Constants -----------------

const int olMailItem        = 0
const int olNoteItem        = 5
const int olByValue         = 1
const int olByReference     = 4
const int olEmbeddedItem    = 5
const int olOLE             = 6

//------------------ Word Symbolic Constants -----------------

const int wdCollapseEnd             = 0
const int wdCollapseStart           = 1
const int wdStory                   = 6
const int wdInLine                  = 0
const int wdTableFormatGrid1        = 16
const int wdBorderTop               = -1
const int wdBorderLeft              = -2
const int wdBorderBottom            = -3
const int wdBorderRight             = -4
const int wdLineStyleNone           = 0
const int wdLineStyleSingle         = 1
const int wdStyleTypeParagraph      = 1
const int wdStyleNormal             = -1
const int maxWordLevel              = 9
const int xlUnderlineStyleSingle    = 2
const int xlUnderlineStyleNone      = -4142
const int wordHeadingStyles[maxWordLevel] = {-2,-3,-4,-5,-6,-7,-8,-9,-10}


//------------------ Word VBA Parameters  --------------------

const string cParamPlacement        = "Placement"
const string cParamRange            = "Range"
const string cParamFileName         = "FileName"
const string cParamAnchor           = "Anchor"
const string cParamLeft             = "Left"
const string cParamLink             = "LinkToFile"
const string cParamSave             = "SaveWithDocument"
const string cParamTop              = "Top"
const string cParamDirection        = "Direction"
const string cParamDisplayAsIcon    = "DisplayAsIcon"
const string cParamUnit             = "Unit"

//------------------------ OLE utility functions -------------

void checkPlatform(string app) {
    if(platform != "WIN32") {
        ack("The " app " interface is only available on PC platforms.")
        halt
    }
}

void closeIfNonNull(OleAutoObj o) {
    if(!null o) {
        oleCloseAutoObject(o)
        o = null
    }
    /* the close perm will automatically set o to null for us */
}

bool oleYesToAll = false 
void checkRes(string res) {
    if((!oleYesToAll) && (res != "")) {
        string answers[] = {"Yes", "Yes to all", "No"}
        int answer = messageBox("OLE method failed: \n" res "\n\nDo you want to continue?", answers, msgWarning)
        
        if (answer == 2) {
            progressStop() 
            halt()
        }
        else if(answer == 1) {
            oleYesToAll = true
        }
    }
}

void makeVisible(OleAutoObj o) {
    bool bIsVisible = false
    // make powerpoint visible
    checkRes(oleGet(o, cPropertyVisible, bIsVisible))
    if (!bIsVisible) {
        checkRes(olePut(o,cPropertyVisible, true))
    }
}

OleAutoObj connectToApp(string appObj, appName) {
    OleAutoObj o = null

    o = oleCreateAutoObject(appObj)
    if(null o) {
        ack "Unable to communicate with " appName "\nIs it installed?"
    }
    return o
}

void disconnectFromApp(OleAutoObj o) {
    closeIfNonNull o
}


// Excel variables
OleAutoObj objExcel = null;
OleAutoObj objWorkbooks = null;
OleAutoObj objWorkbook = null;
OleAutoObj objSheets = null;
OleAutoObj objSheet = null;
OleAutoObj objCell = null;
OleAutoObj objRange = null;
OleAutoObj objColumns = null;
OleAutoObj objColumn = null;
OleAutoObj objRows = null;
OleAutoObj objRow = null;
OleAutoArgs args = null;

bool fileOpen = false;
int oldExcelUndoLimit = 16;

string historyKey = "UndoHistory";
string undoHistoryKey = "";
string excelVersionKey = "HKEY_CLASSES_ROOT\\Excel.Application\\CurVer";
string excelVersion = "";

string excelVersionStrings[4] = {
    "Excel.Application.9",
    "Excel.Application.10",
    "Excel.Application.11",
    "Excel.Application.12"
};

string excelUndoRegistryKeys[4] = {
    "HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\9.0\\Excel\\Options",
    "HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\10.0\\Excel\\Options",
    "HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\11.0\\Excel\\Options",
    "HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\12.0\\Excel\\Options"
};

string trimWhitespace(string s) {
    int first = 0;
    int last = length(s) - 1;

    while(last > 0 && (isspace(s[last]) || s[last] == '\n' || s[last] == '\r' || s[last] == '\t')) {
        last--;
    }
    while(isspace(s[first]) && first < last) {
        first++;
    }
    if(s[first:last] == " ") {
        return("");
    }
    return(s[first:last]);
}

void openLogFile(Stream &logfile, string filename) {
    if(platform() == "WIN32") {
        logfile = write("C:\\Documents and Settings\\" username() "\\Desktop\\" filename ".txt");
    }
    else {
        logfile = write("/home/" username() "/Desktop/" filename ".txt");
    }
    logfile << dateOf(intOf today)"\n";
}

void closeLogFile(Stream &logfile) {
    logfile << dateOf(intOf today)"\n";
    close(logfile);
}

/* 
* Copied from Enhanced Export to Excel written by Galactic Solutions.
* Converts an integer value for a column to the proper string value .
* ex. 1 converts to A
*/
string intToCol(int i) {
    string s = ""
    if ((i>=1) && (i<=256))   {                              // Works for column A -> ZZ
        int d1 = (i-1) / 26;
        int d2 = (i-1) % 26;
        if (d1 > 0) {
            s = charOf( d1-1 + intOf('A')) "";
        }
        s = s charOf( d2 + intOf('A')) "";
    }
    else {
        if (i>256) {
            errorBox("Too many columns");
        }
        else {
            errorBox("Invalid column reference :" i "");
        }
        halt();
    }
    return(s);
}

// Modified version of checkRes from ole.inc file.
bool checkResult(string res)
{
    if(res != "")
    {
        if(confirm("OLE method failed: \n" res "\n\n Do you wish to continue?", msgError))
        {
            return(false);
        }
        else
        {
            progressStop();
            halt();
        }
    }
    return(true);
}

string createExcelStyleAddress(int startRow, int startCol, int endRow, int endCol) {
    return("$" intToCol(startCol) "$" startRow ":" "$" intToCol(endCol) "$" endRow "");
}

/*
* This function closes the current Workbook open without saving.  If you wish to save
* use the saveFile() function listed below.
*/
bool closeExcel() {

    if(fileOpen) {
        closeIfNonNull(objCell);
        closeIfNonNull(objRange);
        closeIfNonNull(objColumns);
        closeIfNonNull(objColumn);
        closeIfNonNull(objRows);
        closeIfNonNull(objRow);
        closeIfNonNull(objSheet);
        closeIfNonNull(objSheets);
        
        clear(args);
        put(args, "SaveChanges", false);    // set it not to save workbook upon close
        
        if(!checkResult(oleMethod(objWorkbook, "Close", args))) {   // close the workbook
            errorBox("Workbook not saved and closed");
            return(false);
        }
        fileOpen = false;
    }

    closeIfNonNull(objWorkbook);
    closeIfNonNull(objWorkbooks);
    
    if(!null(objExcel)) {
        closeIfNonNull(objExcel);   // close Excel
    }

    if(!null(args)) {
        delete(args);
    }

    return(true);
}

/*
* This function saves current workbook.
*/ 
bool saveFile() {
    if(fileOpen) {
        if(!checkResult(oleMethod(objWorkbook, cMethodSave))) {
            errorBox("File not saved");
            return(false);
        }
        return(true);
    }
    errorBox("File not open");
    return(false);
}

/*
* This funciton saves the file as the passed in file name.
*/
bool saveFileAs(string filename) {
    if(fileOpen) {
        clear(args);
        put(args, "FileName", filename);
        
        if(!checkResult(oleMethod(objWorkbook, "SaveAs", args))) {
            errorBox("File SaveAs failed.");
            return(false);
        }
        
        return(true);
    }

    errorBox("File not open");
    return(false);
}

/*
* This function opens the Excel application.
*/
bool openExcel() {
    args = create();
    objExcel = connectToApp(cObjExcelApplication, "Excel");      // connects to Excel
    if(null(objExcel)) {
        return(false);
    }

    return(true);
}

/*
* This function gets the workbooks collection of Excel.
*/
bool getWorkBooks() {
    if(!null(objExcel)) {
        checkResult(oleGet(objExcel, cPropertyWorkbooks, objWorkbooks))   // gets the workbooks collection
        if(null(objWorkbooks)) {
            errorBox("Unable to get workbooks collection");
            return(false);
        }
        
        return(true);
    }

    errorBox("Excel is not currently open");
    return(false);
}

bool getColumns() {
    if(!checkResult(oleGet(objSheet, cPropertyColumns, objColumns))) {
        return(false);
    }

    return(true);
}

bool getColumn(int columnNumber) {
    closeIfNonNull(objColumn);
    
    if(null(objColumns)) {
        if(!getColumns()) {
            errorBox("Excel Columns property not initialized.")
            return(false);
        }
    }

    clear(args);
    put(args, columnNumber);

    if(!checkResult(oleGet(objColumns, cPropertyColumns, args, objColumn))) {
        errorBox("Failed to get handle on Column " intToCol(columnNumber) " in Excel.");
        return(false);
    }

    if(!checkResult(oleMethod(objColumn, cMethodSelect))) {
        errorBox("Failed to select Column " intToCol(columnNumber) "");
        return(false);
    }

    return(true);
}

bool getRows() {
    if(!checkResult(oleGet(objSheet, cPropertyRows, objRows))) {
        return(false);
    }

    return(true);
}

bool getExcelRow(int rowNumber) {
    closeIfNonNull(objRow);

    if(null(objRows)) {
        if(!getRows()) {
            errorBox("Excel Rows Property not initialized.");
            return(false);
        }
    }

    clear(args);
    put(args, rowNumber);

    if(!checkResult(oleGet(objRows, cPropertyRows, args, objRow))) {
        errorBox("Failed to get handle on Row " rowNumber " in Excel.");
        return(false);
    }

    if(!checkResult(oleMethod(objRow, cMethodSelect))) {
        errorBox("Failed to select Row " rowNumber "");
        return(false);
    }

    return(true);
}

/*
* This function gets the active workbook in Excel.
*/ 
bool getActiveWorkBook() {
    if(!null(objExcel)) {
        checkResult(oleGet(objExcel, cPropertyActiveWorkbook, objWorkbook));
        if(null(objWorkbook)) {
            errorBox("Failed to get active workbook.");
            return(false);
        }
        return(true);
    }

    errorBox("Excel application not initialized.");
    return(false);
}

/*
* Adds a workbook to the workbooks collection and sets the workbook object to this
* new workbook using the getActiveWorkBook function.  Then it saves the new workbook
* according to the filename provided.
*/
bool addWorkBook() {
    if(!null(objWorkbooks)) {
        if(!checkResult(oleMethod(objWorkbooks, cMethodAdd))) {
            errorBox("Add workbook failed.");
            return(false);
        }
        
        getActiveWorkBook();
        fileOpen = true;
        return(true);
    }
    errorBox("Workbooks collection not initialized.  Check to make sure your code gets it.");
    return(false);
}

/*
* Adds a workbook to the workbooks collection and sets the workbook object to this
* new workbook using the getActiveWorkBook function.  Then it saves the new workbook
* according to the filename provided.
*/
bool addWorkBook(string filename) {
    if(!null(objWorkbooks)) {
        if(!checkResult(oleMethod(objWorkbooks, cMethodAdd))) {
            errorBox("Add workbook failed.");
            return(false);
        }
        
        getActiveWorkBook();
        fileOpen = true;
        saveFileAs(filename);
        
        return(true);
    }

    errorBox("Workbooks collection not initialized.  Check to make sure your code gets it.");
    return(false);
}

/*
* This funciton opens the existing workbook according to the passed in file name.
*/
bool openWorkBook(string filename) {
    if(!null(objWorkbooks)) {
        clear(args);
        put(args, cParamFileName, filename);
        put(args, cPropertyUpdateLinks, true)
        
        if(!checkResult(oleMethod(objWorkbooks, "Open", args))) {      // opens the file
            errorBox("Failed to open file");
            return(false);
        }
        
        fileOpen = true;
        return getActiveWorkBook();
    }

    errorBox("Workbooks collection not initialized.  Check to make sure your code gets it.");
    return(false);
}

/*
* This function gets the number of sheets in the workbook.  getSheets must have been
* called prior to this.  Returns -1 if the function fails.
*/
int getNumberOfSheets() {
    int numSheets = 0;

    if(!null(objSheets)) {
        if(!checkResult(oleGet(objSheets, cPropertyCount, numSheets))) {   // gets the number of sheets
            errorBox("Unable to get number of Sheets");
            return(-1);
        }
        
        return(numSheets);
    }

    errorBox("Null Sheets collection.  Check that getSheets() has been called.");
    return(-1);
}

/*
* This function gets the name of the sheet currently assigned to the sheet variable.
*/
string getSheetName() {
    string sheetName;

    if(!null(objSheet)) {
        if(!checkResult(oleGet(objSheet, cPropertyName, sheetName))) {      // gets the name of the sheet
            errorBox("Failed to get sheet name");
            return("error");
        }
        
        return(sheetName);
    }

    errorBox("Null sheet object.  Check that getSheet or getActiveSheet has been called.");
    return("error");
}

/*
* This function sets the name of the sheet currently assigned to the sheet variable.
*/
bool setSheetName(string sheetName) {
    if(!null(objSheet)) {
        if(!checkResult(olePut(objSheet, cPropertyName, sheetName))) {
            errorBox("Failed to set the sheet name");
            return(false);
        }
        else {
            return(true);
        }
    }

    errorBox("Null sheet object.  Check that getSheet(int) or getActiveSheet() has been called.");
    return(false);
}

/*
* This function gets the active sheet of the workbook.
*/
bool getActiveSheet() {
    if(!null(objWorkbook)) {
        checkResult(oleGet(objWorkbook, cPropertyActiveSheet, objSheet));   // gets the active sheet
        
        if(null(objSheet)) {
            errorBox("Unable to get active sheet");
            return(false);
        }
        
        return(true);
    }

    errorBox("Null workbook object.  Check that getWorkbook() has been called.");
    return(false);
}

/*
* This function gets the Sheets open in Excel.
*/
bool getSheets() {
    if(!null(objExcel) && !null(objWorkbook)) {      // must have Excel and a workbook open.
        checkResult(oleGet(objExcel, cPropertySheets, objSheets));   // get the sheets property.
        
        if(null(objSheets)) {
            errorBox("Unable to get Sheets for file");
            return(false);
        }
        
        return(true);
    }

    errorBox("Excel not open or no workbook loaded.");
    return(false);
}

/*
* This function returns a sheet from the list of sheets in the sheets property.
* It can be useful if trying to loop through the sheets in a workbook.  Remember 
* to check the name of the sheet.  If it contains the word "Module", skip it.
* These are not actual sheets.
*/
bool getSheet(int sheetNumber) {
    if(!null(objWorkbook)) {
        clear(args);
        put(args, sheetNumber);
        checkResult(oleGet(objWorkbook, cPropertySheets, args, objSheet));
        
        if(null(objSheet)) {
            errorBox("Unable to get sheet " sheetNumber " for file");
            return(false);
        }
        
        checkResult(oleMethod(objSheet, cMethodActivate));
        return(true);
    }

    errorBox("Null Workbook handle.  Check that openWorkbook() has been called.");
    return(false);
}

/*
* Gets a handle on a specific cell based on the row and column number.
* Remember, the minimum row and column number in Excel is 1, not 0.  The
* maximum column number is 256.  Office 97 and 2000 have a limit of 65536
* rows.  Office XP and 2003 can go beyond this.  However for DOORS this 
* should not be an issue.
*/
bool getCell(int rowNum, int col) {
    closeIfNonNull(objCell);
    clear(args);
    put(args, intToCol(col) rowNum "");

    checkResult(oleGet(objSheet, cMethodRange, args, objCell));

    if(null(objCell)) {
        return(false);
    }

    checkResult(oleMethod(objCell, cMethodSelect));

    return(true);
}

/*
* Gets a handle on a specific cell based on a cell location such as A9.
* Remember, the minimum row and column number in Excel is 1, not 0.  The
* maximum column number is 256.  Office 97 and 2000 have a limit of 65536
* rows.  Office XP and 2003 can go beyond this.  However for DOORS this 
* should not be an issue.
*/
bool getCell(string loc) {
    closeIfNonNull(objCell);      // close any open handle to a cell
    clear(args);                   // clear any current arguments
    put(args, loc "");             // loc is in the format of A9 or B37

    checkResult(oleGet(objSheet, cMethodRange, args, objCell));      // get a handle to the cell

    if(null(objCell)) {
        errorBox("Unable to get cell object");
        return(false);
    }

    checkResult(oleMethod(objCell, cMethodSelect));

    return(true);
}

/*
* This function gets the Value Property of the cell at the passed in row and
* column numbers.  It will get a handle to the cell before trying to get the
* value.
*/ 
string getCellValue(int rowNum, int col) {
    string s = "";
    getCell(rowNum, col);   // gets a handle on the desired cell.

    // if the handle of a cell was returned, get the value
    if(!null(objCell)) {
        if(!checkResult(oleGet(objCell, cPropertyValue, s))) {
            errorBox("Failed to get cell " intToCol(col) "" rowNum " value");
            return("error");
        }
    }
    else {   // otherwise returns error as the result
        return("error");
    }

    return(s);
}

/*
* This function gets the Value property of the cell at the passed in cell
* location.  It will get a handle to the cell before trying to 
* get the value.
*/ 
string getCellValue(string loc) {
    string s = "";
    getCell(loc);   // gets a handle on the desired cell.

    // if the handle of a cell was returned, get the value
    if(!null(objCell)) {
        if(!checkResult(oleGet(objCell, cPropertyValue, s))) {
            errorBox("Failed to get cell " loc " value");
            return("error");
        }
    }
    else {   // otherwise returns error as the result
        return("error");
    }

    return(s);
}

/*
* This function gets the Text Property of the cell at the passed in row and
* column numbers.  It will get a handle to the cell before trying 
* to get the value.
*/ 
string getCellText(int rowNum, int col) {
    string s = "";
    getCell(rowNum, col);   // gets a handle on the desired cell.

    // if the handle of a cell was returned, get the value
    if(!null(objCell)) {
        if(!checkResult(oleGet(objCell, cPropertyText, s))) {
            errorBox("Failed to get cell " intToCol(col) "" rowNum " value");
            return("error");
        }
    }
    else {   // otherwise returns error as the result
        return("error");
    }

    return(s);
}

/*
* This function gets the Text property of the cell at the passed in cell
* location.  It will get a handle to the cell before trying to 
* get the value.
*/ 
string getCellText(string loc) {
    string s = "";
    getCell(loc);   // gets a handle on the desired cell.

    // if the handle of a cell was returned, get the value
    if(!null(objCell)) {
        if(!checkResult(oleGet(objCell, cPropertyText, s))) {
            errorBox("Failed to get cell " loc " value");
            return("error");
        }
    }
    else {   // otherwise returns error as the result
        return("error");
    }

    return(s);
}

/*
* This function sets the value of the cell at the passed in row and
* column numbers.  It will get a handle to the cell before trying 
* to set the value.
*/ 
bool setCellValue(int rowNum, int col, string s) {
    getCell(rowNum, col);         // get the cell handle

    // if the cell handle isn't null, put the value in the cell
    if(!null(objCell)) {
        if(!checkResult(olePut(objCell, cPropertyValue, s))) {
            errorBox("Failed to set cell " intToCol(col) "" rowNum " value");
            return(false);
        }
        
        return(true);
    }

    return(false);            // otherwise false;
}

/*
* This function sets the value of the cell at the passed in row and
* column numbers.  It will get a handle to the cell before trying 
* to set the value.
*/ 
bool setCellValue(string loc, string s) {
    getCell(loc);            // get the cell handle

    // if the cell handle isn't null, put the value in the cell
    if(!null objCell) {
        if(!checkResult(olePut(objCell, cPropertyValue, s))) {
            errorBox("Failed to set cell " loc " value");
            return(false);
        }
        
        return(true);
    }

    return(false);            // otherwise return(false);
}

bool setCellFormula(int rowNum, int col, string function) {
    getCell(rowNum, col);

    if(!null(objCell)) {
        if(!checkResult(olePut(objCell, cPropertyFormula, function))) {
            errorBox("Failed to set cell " intToCol(col) "" rowNum " function");
            return(false);
        }
        
        return(true);
    }

    errorBox("Null Cell reference in call to setCellFunction");
    return(false);
}

bool setCellFormula(string loc, string function) {
    getCell(loc);

    if(!null(objCell)) {
        if(!checkResult(olePut(objCell, cPropertyFormula, function))) {
            errorBox("Failed to set cell " loc " formula");
            return(false);
        }
        
        return(true);
    }

    errorBox("Null Cell reference in call to setCellFunction");
    return(false);
}

/*
* Gets a handle to the passed in cell range.
*/
bool getCellRange(int startRow, int startCol, int endRow, int endCol) {
    Buffer cellRange = create();

    closeIfNonNull(objRange);
    cellRange = intToCol(startCol) startRow ":" intToCol(endCol) endRow "";

    clear(args);
    put(args, tempStringOf(cellRange));

    if(!checkResult(oleGet(objSheet, cPropertyRange, args, objRange))) {
        errorBox("Failed to get cell range: " tempStringOf(cellRange) "");
        delete(cellRange);
        return(false);
    }

    checkResult(oleMethod(objRange, cMethodSelect));
    return(true);
}

/*
* Sets the current cell's text to be wrapped.
*/
bool setCellWrapText(bool wrapText) {
    if(!null(objCell)) {
        if(!checkResult(olePut(objCell, cPropertyWrapText, wrapText))) {
            errorBox("Failed to set property WrapText for cell");
            return(false);
        }
    }
    else {
        errorBox("Null cell object when attempting to set WrapText");
        return(false);
    }

    return(true);
}

/*
* Gets a handle to the passed in cell range and sets all cells in that range to have their
* text be wrapped.
*/
bool setRangeWrapText(int startRow, int startCol, int endRow, int endCol, bool wrapText) {
    if(getCellRange(startRow, startCol, endRow, endCol)) {
        if(!checkResult(olePut(objRange, cPropertyWrapText, wrapText))) {
            errorBox("Failed to set property WrapText for cell");
            return(false);
        }
        
        return(true);
    }

    return(false);
}

/*
* Merges all cells in the desired range.
*/
bool setRangeMergeCells(int startRow, int startCol, int endRow, int endCol, bool mergeCells) {
    if(getCellRange(startRow, startCol, endRow, endCol)) {
        if(!checkResult(olePut(objRange, cPropertyMergeCells, mergeCells))) {
            errorBox("Failed to set property MergeCells for range");
            return(false);
        }
        
        return(true);
    }

    return(false);
}

/* Sets the specified column to the specified width.  
* One thing to keep in mind is that while this function specifies the columns with
* in characters, DOORS stores its column width as the number of pixels on the screen.
* 
* If you wish to set the Excel column width to somewhere near the width that it had in
* DOORS, I have found that using width(Column c)/6 works fairly well.
*/
bool setExcelColumnWidth(int columnNumber, int columnWidth) {
    // get a handle to the column.
    if(!getColumn(columnNumber)) {
        return(false);
    }

    // set the column width
    if(!checkResult(olePut(objColumn, cPropertyColumnWidth, columnWidth))) {
        errorBox("Failed to set column " intToCol(columnNumber) "s width to " columnWidth "");
        return(false);
    }

    return(true);
}

/*
* Sets the specified row to the specified width.  
*/
bool setExcelRowWidth(int rowNumber, int rowHeight) {
    // get a handle to the row.
    if(!getExcelRow(rowNumber)) {
        return(false);
    }

    // set the column width
    if(!checkResult(olePut(objRow, cPropertyRowHeight, rowHeight))) {
        errorBox("Failed to set row " rowNumber "s width to " rowHeight "");
        return(false);
    }

    return(true);
}

/*
* Sets the cell's Background color to the specified color.
* Referenced Enhanced Export to Excel for this function.
*/
bool setCellBackgroundColor(int rowNum, int col, int colorValue) {
    OleAutoObj objInterior = null;
    bool successful = false;

    // get a handle to the cell
    if(!getCell(rowNum, col)) {
        errorBox("Null cell reference");
    }
    else {
        if(!checkResult(oleGet(objCell, cPropertyInterior, objInterior))) {
            errorBox("Failed to get handle on cell's Interior Property");
        }
        else {
            if(!checkResult(olePut(objInterior, cPropertyColorIndex, colorValue))) {
                errorBox("Failed to set cell's Color Property");
            }
            else {
                successful = true;
            }
        }
    }

    closeIfNonNull(objInterior);

    return(successful);
}

/*
* Sets the desired range's Background color to the specified color.
*/
bool setRangeBackgroundColor(int startRow, int startCol, int endRow, int endCol, int colorValue) {
    OleAutoObj objInterior = null;
    OleAutoObj objSelection = null;
    bool successful = false;

    if(!getCellRange(startRow, startCol, endRow, endCol)) {
        errorBox("Failed to get handle on Range " intToCol(startCol) startRow ":" intToCol(endCol) endRow "");
    }
    else {
        if(!checkResult(oleGet(objExcel, cPropertySelection, objSelection))) {
            errorBox("Failed to get handle on Selection");
        }
        else {
            if(!checkResult(oleGet(objSelection, cPropertyInterior, objInterior))) {
                errorBox("Failed to get handle on cell's Interior Property");
            }
            else {
                if(!checkResult(olePut(objInterior, cPropertyColorIndex, colorValue))) {
                    errorBox("Failed to set range's Color Property");
                }
                else {
                    successful = true;
                }
            }
        }
    }

    closeIfNonNull(objInterior);
    closeIfNonNull(objSelection);

    return(successful);
}

/*
* Gets the cell's background color.
*/
int getCellBackgroundColor(int rowNum, int col) {
    OleAutoObj objInterior = null;
    int colorValue = 0;
    int returnVal = -1;

    // get a handle to the cell
    if(!getCell(rowNum, col)) {
        errorBox("Null cell reference");
    }
    else {
        if(!checkResult(oleGet(objCell, cPropertyInterior, objInterior))) {
            errorBox("Failed to get handle on cell's Interior Property");
        }
        else {
            if(!checkResult(oleGet(objInterior, cPropertyColorIndex, colorValue))) {
                errorBox("Failed to get cell's Color Property");
            }
            else {
                returnVal = colorValue;
            }
        }
    }

    closeIfNonNull(objInterior);

    return(returnVal);
}

/*
* Sets the cell formatting for the given cell.
*
* See ole.inc for the valid values for format. Listed under Microsoft
* Excel Number Formats.
*/
bool setCellNumberFormat(int rowNum, int col, string format) {
    OleAutoObj objSelection = null;
    bool successful = false;

    if(!getCell(rowNum, col)) {
        errorBox("Null cell reference");
    }
    else {
        if(!checkResult(oleMethod(objCell, cMethodSelect)))   {   // selects the cell
            errorBox("Failed to select cell " intToCol(col) "" rowNum "");
        }
        else {
            if(!checkResult(oleGet(objExcel, cPropertySelection, objSelection))) {
                errorBox("Failed to get handle on Selection");
            }
            else {
                if(!checkResult(olePut(objSelection, cPropertyNumberFormat, format))) {
                    errorBox("Failed to set Number Format of selection");
                }
                else {
                    successful = true;
                }
            }
        }
    }

    closeIfNonNull(objSelection);

    return(successful);
}

/*
* Gets the current number format of the selected cell.
*/
string getCellNumberFormat(int rowNum, int col) {
    string numberFormat = "error";

    OleAutoObj objSelection = null;

    if(!getCell(rowNum, col)) {
        errorBox("Null cell reference.")
    }
    else {
        if(!checkResult(oleMethod(objCell, cMethodSelect)))   {   // selects the cell
            errorBox("Failed to select cell " intToCol(col) "" rowNum "");
        }
        else {
            if(!checkResult(oleGet(objExcel, cPropertySelection, objSelection))) {
                errorBox("Failed to get handle on Selection");
            }
            else {
                if(!checkResult(oleGet(objSelection, cPropertyNumberFormat, numberFormat))) {
                    errorBox("Failed to get Number format of " intToCol(col) "" rowNum "");
                }
            }
        }
    }

    closeIfNonNull(objSelection);

    return(numberFormat);
}

/*
* Sets the cell formatting for the given range.
*
* See ole.inc for the valid values for format. Listed under Microsoft
* Excel Number Formats.
*/
bool setRangeNumberFormat(int startRow, int startCol, int endRow, int endCol, string format) {
    OleAutoObj objSelection = null;
    bool successful = false;

    if(!getCellRange(startRow, startCol, endRow, endCol)) {
        errorBox("Failed to get handle on Range " intToCol(startCol) startRow ":" intToCol(endCol) endRow "");
    }
    else {
        if(!checkResult(oleMethod(objRange, cMethodSelect)))   {   // selects the cell
            errorBox("Failed to select Range " intToCol(startCol) startRow ":" intToCol(endCol) endRow "");
        }
        else {
            if(!checkResult(oleGet(objExcel, cPropertySelection, objSelection))) {
                errorBox("Failed to get handle on Selection");
            }
            else {
                if(!checkResult(olePut(objSelection, cPropertyNumberFormat, format))) {
                    errorBox("Failed to set Number Format of selection");
                }
                else {
                    successful = true;
                }
            }
        }
    }

    closeIfNonNull(objSelection);
    return(successful);
}

/*
* Sets the current cells text to the specified color.
*/
bool setCellTextColor(int rowNum, int col, int colorValue) {
    OleAutoObj objFont = null;

    // get a handle to the cell
    if(!getCell(rowNum, col)) {
        errorBox("Null cell reference");
        return(false);
    }

    // get a handle on the cells font property.
    if(!checkResult(oleGet(objCell, cPropertyFont, objFont))) {
        errorBox("Failed to get a handle to the cells Font property");
        return(false);
    }

    // set the cell's text color.
    if(!checkResult(olePut(objFont, cPropertyColorIndex, colorValue))) {
        errorBox("Failed to set the cells font color");
        return(false);
    }

    return(true);
}

/*
* This function copies the contents of the passed in Excel cell to the clipboard
* preserving any rich text formatting it may have.
*/
string copyRichTextFromCell(string loc) {
    if(!getCell(loc)) {
        return("error");
    }

    if(!checkResult(oleMethod(objCell, cMethodSelect)))   {  // selects the cell
        return("error");
    }

    if(!checkResult(oleMethod(objCell, cMethodCopy))) {      // copies the contents of the cell to the clipboard
        return("error");
    }

    return stringOf(richClip)
}

bool checkSheetFilterStatus() {
    bool filtersOn = false;

    if(!null(objSheet)) {
        checkResult(oleGet(objSheet, "FilterMode", filtersOn));
    }
    else {
        errorBox("Null Sheet Object.  Make sure that getSheet or getActiveSheet has been called.");
        return(false);
    }

    if(filtersOn) {
        return(true);
    }
    else {
        return(false);
    }
}

bool turnSheetFiltersOff() {
    if(!checkResult(oleMethod(objSheet, "ShowAllData"))) {
        return(false);
    }
    else {
        return(true);
    }
}

/*
* Used to set the print area of the current Excel sheet to the passed in row and
* column locations.
*/
bool setPrintArea(int startRow, int startCol, int endRow, int endCol) {
    OleAutoObj objPageSetup = null;
    string printArea = createExcelStyleAddress(startRow, startCol, endRow, endCol);
    bool successful = false;

    if(null(objSheet)) {
        errorBox("Null sheet object.  Please call getSheet() or getActiveSheet() first");
    }
    else {
        if(!checkResult(oleGet(objSheet, cPropertyPageSetup, objPageSetup))) {
            errorBox("Failed to get handle on Excel Page Setup.");
        }
        else {
            if(!checkResult(olePut(objPageSetup, cPropertyPrintArea, printArea))) {
                errorBox("Failed to set print area to " printArea "");
            }
            else {
                successful = true;
            }
        }
    }

    closeIfNonNull(objPageSetup);
    return(successful);
}

/*
* Used to get the print area of the current sheet in Excel.
* A note to keep in mind.  If the print area is not set, this function
* returns an empty string.  So be sure to test for it or if using it,
* make sure that your print area is set (this has to be done manually in 
* Excel).  This can be done with the above function.
*/
string getPrintArea() {
    OleAutoObj objPageSetup = null;
    string printArea = "";
    string returnValue = "error";

    if(null(objSheet)) {
        errorBox("Null sheet object.  Please call getSheet() or getActiveSheet() first");
    }
    else {
        if(!checkResult(oleGet(objSheet, cPropertyPageSetup, objPageSetup))) {
            errorBox("Failed to get handle on Excel Page Setup.");
        }
        else {
            if(!checkResult(oleGet(objPageSetup, cPropertyPrintArea, printArea))) {
                errorBox("Failed to get Excel Print Area from Page Setup");
            }
            else {
                returnValue = printArea;
            }
        }
    }

    closeIfNonNull(objPageSetup);
    return(returnValue);
}

bool runExcelMacro(string macroName) {
    clear(args);
    put(args, macroName);

    if(null(objExcel)) {
        errorBox("Excel application not initialized.");
        return(false);
    }

    if(!checkResult(oleMethod(objExcel, cMethodRun, args))) {
        errorBox("Macro " macroName " failed to run.  Please check that it exists in the Excel file.");
        return(false);
    }

    return(true);
}

/*
* Sets the horizontal alignment of text for the selected range.
*
* Valid values for hAlignment are xlHAlignCenter, xlHAlignLeft, xlHAlignRight,
* and xlHAlignJustify
*/
bool setRangeHorizontalAlignment(int startRow, int startCol, int endRow, int endCol, int hAlignment) {
    if(!getCellRange(startRow, startCol, endRow, endCol)) {
        errorBox("Failed to get handle on Range " intToCol(startCol) startRow ":" intToCol(endCol) endRow "");
        return(false);
    }

    if(!checkResult(olePut(objRange, cPropertyHorizontalAlignment, hAlignment))) {
        errorBox("Failed to set Horizontal Alignment on Range " intToCol(startCol) startRow ":" intToCol(endCol) endRow "");
        return(false);
    }

    return(true);
}

/*
* Sets the vertical alignment of text for the selected range.
*
* Valid values for vAlignment are xlVAlignTop, xlVAlignBottom, xlVAlignCenter,
* and xlVAlignJustify.
*/
bool setRangeVerticalAlignment(int startRow, int startCol, int endRow, int endCol, int vAlignment) {
    if(!getCellRange(startRow, startCol, endRow, endCol)) {
        errorBox("Failed to get handle on Range " intToCol(startCol) startRow ":" intToCol(endCol) endRow "");
        return(false);
    }

    if(!checkResult(olePut(objRange, cPropertyVerticalAlignment, vAlignment))) {
        errorBox("Failed to set Vertical Alignment on Range " intToCol(startCol) startRow ":" intToCol(endCol) endRow "");
        return(false);
    }

    return(true);
}

/*
* Sets the border
*/

/*
*  These following functions were copied from Galactic Solutions 
*  Enhanced Export To Excel and modified to work with my code.
*/

void setFontSymbolOLE(OleAutoObj objExcelFont) {
    checkResult(olePut(objExcelFont, cPropertyName, cPropertySymbol));
}

void setFontRegularOLE(OleAutoObj objExcelFont) {
    checkResult( olePut( objExcelFont, cPropertyFontStyle, cPropertyRegular));
}

void setFontBoldOLE(OleAutoObj objExcelFont) {
    checkResult(olePut(objExcelFont, cPropertyFontStyle, cPropertyBold));
}

void setFontItalicOLE(OleAutoObj objExcelFont) {
    checkResult(olePut(objExcelFont, cPropertyFontStyle, cPropertyItalic));
}

void setFontBoldItalicOLE(OleAutoObj objExcelFont) {
    checkResult(olePut(objExcelFont, cPropertyFontStyle, cPropertyBoldItalic));
}

void setFontUnderlineSingleOLE(OleAutoObj objExcelFont) {
    checkResult(olePut(objExcelFont, cPropertyUnderline, xlUnderlineStyleSingle));
}

void setFontUnderlineNoneOLE(OleAutoObj objExcelFont) {
    checkResult(olePut(objExcelFont, cPropertyUnderline, xlUnderlineStyleNone));
}

void setFontStrikeThroughOLE(OleAutoObj objExcelFont) {
    checkResult(olePut(objExcelFont, cPropertyStrikeThrough, true));
}

void setFontSuperscriptOLE(OleAutoObj objExcelFont) {
    checkResult(olePut(objExcelFont, cPropertySuperscript, true));
}

void setFontSubscriptOLE(OleAutoObj objExcelFont) {
    checkResult(olePut(objExcelFont, cPropertySubscript, true));
}


OleAutoObj getCharactersOLE(OleAutoObj objExcelCell, int iStart, int iLength) {
    OleAutoObj objExcelChars = null;

    clear(args);
    put(args, cParamStart, iStart);
    put(args, cParamLength, iLength);

    checkResult(oleGet(objExcelCell, cPropertyCharacters, args, objExcelChars));

    return(objExcelChars);
}

OleAutoObj getFontOLE(OleAutoObj objExcelCell, int iStart, int iLength) {
    OleAutoObj objExcelChars = getCharactersOLE(objExcelCell, iStart, iLength);
    OleAutoObj objExcelFont = null;

    checkResult(oleGet(objExcelChars, cPropertyFont, objExcelFont));

    return(objExcelFont);
}

string stripRichText(string rts) {
    Buffer  plainText = create();
    RichText rt;

    for rt in rts do {
        plainText += rt.text;
        
        if ((rt.newline) && !(rt.last)) {
            plainText += "\n";
        }
    }

    return(trimWhitespace(stringOf(plainText)));
}

bool setCellRichText(int rowNum, int col, string richTextContents) {
    if(!setCellValue(rowNum, col, stripRichText(richTextContents))) {
        return(false);
    }

    OleAutoObj objExcelFont = null;
    int characterCount = 1;
    RichText rt;

    for rt in richTextContents do {
        objExcelFont = getFontOLE(objCell, characterCount, length(rt.text));

        if(rt.charset == charsetSymbol) {
            setFontSymbolOLE(objExcelFont);
        }

        if(!(rt.bold ) && !(rt.italic)) {
            setFontRegularOLE(objExcelFont);
        }
        
        if((rt.bold) && !(rt.italic)) {
            setFontBoldOLE(objExcelFont);
        }
        
        if(!(rt.bold) &&  (rt.italic)) {
            setFontItalicOLE(objExcelFont);
        }
        
        if((rt.bold) &&  (rt.italic)) {
            setFontBoldItalicOLE(objExcelFont);
        }

        if(rt.underline) {
            setFontUnderlineSingleOLE(objExcelFont);
        }
        else {
            setFontUnderlineNoneOLE(objExcelFont);
        }

        if(rt.strikethru) {
            setFontStrikeThroughOLE(objExcelFont);
        }
        
        if(rt.superscript) {
            setFontSuperscriptOLE(objExcelFont);
        }
        
        if(rt.subscript) {
            setFontSubscriptOLE(objExcelFont);
        }
        
        characterCount += length rt.text;
        
        if(rt.newline) {
            characterCount++;
        }
        
        closeIfNonNull(objExcelFont);
    }
    
    return(true);
}

/*****************End Functions from Enhanced Export to Excel***************************/

/****************************Speed Improvement Functions********************************/
bool turnOffExcelUndoFeature() {
    string undoHistory = "";
    int i;

    excelVersion = getRegistry(excelVersionKey, null);

    for(i = 0; i < sizeof(excelVersionStrings); i++) {
        if(excelVersion == excelVersionStrings[i]) {
            undoHistoryKey = excelUndoRegistryKeys[i];
            break;
        }
    }

    if(null(undoHistoryKey)) {
        errorBox("Unable to detect Excel version.  Please ensure Excel 2000-2007 is installed");
        return(false);
    }

    undoHistory = getRegistry(undoHistoryKey, historyKey);

    if(!null(undoHistory)) {
        oldExcelUndoLimit = intOf(undoHistory);
    }

    setRegistry(undoHistoryKey, historyKey, 0);
    return(true);
}

bool restoreExcelUndoFeature() {
    if(undoHistoryKey != "") {
        setRegistry(undoHistoryKey, historyKey, oldExcelUndoLimit);
        return(true);
    }

    return(false);
}

bool toggleAutoComplete(bool enable) {
    if(null(objExcel)) {
        return(false);
    }

    if(!checkResult(olePut(objExcel, cPropertyEnableAutoComplete, enable))) {
        errorBox("Failed to set Excel Auto Complete to " enable "");
        return(false);
    }

    return(true);
}

bool toggleAutoRecover(bool enable) {
    OleAutoObj autoRecover = null;

    if(null(objExcel)) {
        return(false);
    }

    if(!checkResult(oleGet(objExcel, cPropertyAutoRecover, autoRecover))) {
        errorBox("Failed to get handle on Excel Auto Recover property");
        return(false);
    }
    
    if(!checkResult(olePut(autoRecover, cPropertyEnabled, enable))) {
        errorBox("Failed to set Excel Auto Recover to " enable "");
        return(false);
    }

    return(true);
}

bool toggleErrorChecking(bool enable) {
    OleAutoObj errorChecking = null;

    if(null(objExcel)) {
        return(false);
    }

    if(!checkResult(oleGet(objExcel, cPropertyErrorCheckingOptions, errorChecking))) {
        errorBox("Failed to get handle on Excel ErrorCheckingOptions property");
        return(false);
    }
    
    if(!checkResult(olePut(errorChecking, cPropertyBackgroundChecking, enable))) {
        errorBox("Failed to set Excel Error Checking to " enable "");
        return(false);
    }

    return(true);
}

bool toggleAutoCorrectOptions(bool enable) {
    OleAutoObj autoCorrectOptions = null;

    if(null(objExcel)) {
        return(false);
    }

    if(!checkResult(oleGet(objExcel, cPropertyAutoCorrect, autoCorrectOptions))) {
        errorBox("Failed to get handle on Excel Auto Correct Options property");
        return(false);
    }

    if(!checkResult(olePut(autoCorrectOptions, cPropertyTwoInitialCapitals, enable))) {
        errorBox("Failed to set Excel Auto Correct Option - " cPropertyTwoInitialCapitals " to " enable"");
        return(false);
    }

    if(!checkResult(olePut(autoCorrectOptions, cPropertyCorrectSentenceCap, enable))) {
        errorBox("Failed to set Excel Auto Correct Option - " cPropertyCorrectSentenceCap " to " enable"");
        return(false);
    }

    if(!checkResult(olePut(autoCorrectOptions, cPropertyCapitalizeNamesofDays, enable))) {
        errorBox("Failed to set Excel Auto Correct Option - " cPropertyCapitalizeNamesofDays " to " enable"");
        return(false);
    }

    if(!checkResult(olePut(autoCorrectOptions, cPropertyCorrectCapsLock, enable))) {
        errorBox("Failed to set Excel Auto Correct Option - " cPropertyCorrectCapsLock " to " enable"");
        return(false);
    }

    if(!checkResult(olePut(autoCorrectOptions, cPropertyReplaceText, enable))) {
        errorBox("Failed to set Excel Auto Correct Option - " cPropertyReplaceText " to " enable"");
        return(false);
    }

    if(!checkResult(olePut(autoCorrectOptions, cPropertyDisplayAutoCorrectOptions, enable))) {
        errorBox("Failed to set Excel Auto Correct Option - " cPropertyDisplayAutoCorrectOptions " to " enable"");
        return(false);
    }
    
    return(true);
}

bool toggleSpeedImprovements(bool enable) {
    enable = !enable;

    if(!toggleAutoComplete(enable)) {
        return(false);
    }

    // only supported in Office XP/2002 and later
    if(excelVersion != excelVersionStrings[0]) {
        if(!toggleAutoRecover(enable)) {
            return(false);
        }

        if(!toggleErrorChecking(enable)) {
            return(false);
        }
        
        if(!toggleAutoCorrectOptions(enable)) {
            return(false);
        }
    }

    return(true);
}