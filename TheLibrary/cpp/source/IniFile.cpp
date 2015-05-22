//#include "inifile.h"
#include "the.h"
#include <stdlib.h>
#include <windows.h>
#include <string.h>
#include <string>
#include <algorithm>

using namespace std;


// STATIC/GLOBAL
vector <TheIniFileBuffer*> TheIniFile::_iniDB;		// all the _iniDB.


//==============================================================================
//==============================================================================
TheIniFile::TheIniFile() : _currentIniFilePtr(0){
}

//==============================================================================
//==============================================================================
TheIniFile::~TheIniFile(){
	close();
}

//==============================================================================
//==============================================================================
bool TheIniFile::open (const string& file) {
	int result;
	_fileName= file;
	_FileSystem->makeFullPath(_fileName);


	// if another ini file is opened, close that one first.
	clear();

	//check if it is already loaded in the memory.
	int i=0;
	bool found=false;
	while (i <_iniDB.size() && !found ) {
		if (_iniDB[i]->getFileName()==_fileName) {
			found=true;
			_currentIniFilePtr=_iniDB[i];
			return true;	// Use the one in the mem.
		}
		i++;
	}

	// add a new INI file entry in mem
	_currentIniFilePtr=new TheIniFileBuffer;
	_currentIniFilePtr->setFileName(_fileName);
	_iniDB.push_back(_currentIniFilePtr);
	
	// open physical file
	result = _file.open(_fileName);
	
	// if file is not found, create it.
	if (!result) { 
		return false;
	}

	// read physical file into memory
	string line;
	while (!_file.eof()) {
		_file.readLine(line);
		_currentIniFilePtr->_buffer.push_back(line.c_str());
	}
	
	// close physical file
	_file.close();
	return true;
}

//==============================================================================
//==============================================================================
bool TheIniFile::openConfig() {
	string appPath,
		drive,dir,fileName,ext;	// component
	_FileSystem->getAppPath(appPath);
	_FileSystem->splitPath(appPath,drive,dir,fileName,ext);
	_FileSystem->makePath(appPath,drive,dir,fileName,".ini");
	return open(appPath);
}
//==============================================================================
//==============================================================================
bool TheIniFile::openSystem(const string& filename) {
	string systemPath;
	TheOS::getSystemPath(systemPath);
	systemPath=systemPath+filename;		// "c:\windows\myfile.ini"
	return open(systemPath);
}
//==============================================================================
//==============================================================================
void TheIniFile::clearAll() {
	int i=0;
	int size=_iniDB.size();
	while (i <size  ) {
		_iniDB[i]->clear();
		TheIniFileBuffer* tempPtr=_iniDB[i];
		delete (tempPtr);		// delete individual items
		i++;
	}
	_iniDB.clear();		// erase all DB entry

}
//==============================================================================
//==============================================================================
bool TheIniFile::close() {
	flush();		// flush if needed.

	return clear();
}

//==============================================================================
//==============================================================================
bool TheIniFile::clear() {
	if (_currentIniFilePtr) {
		_currentIniFilePtr->clear();
		_iniDB.erase (find(_iniDB.begin(),_iniDB.end(),_currentIniFilePtr));
		delete (_currentIniFilePtr);
		_currentIniFilePtr=0;
		return true;
	}
	return false;
}

//==============================================================================
//==============================================================================
void TheIniFile::readStr(string& str, const string& group, const string& key, 
							const string& defaultStr) {

	assert(_currentIniFilePtr);
	int i=findSectionKey(group,key);
	if (i<0) {
		str=defaultStr;	// not found!
		return;
	}
	string dummyKey;
	extractKeyValue(dummyKey,str,i);
	if (str=="") {
		str=defaultStr;		// empty string
	}
}


//==============================================================================
//==============================================================================
int TheIniFile::findSection(const string& section) {
	int i=0;
	string text;
	int size=_currentIniFilePtr->_buffer.size();
	while (i<size ) {
		text=_currentIniFilePtr->_buffer.get(i);
		TheString::removeLeadingWhiteSpace(text);
		if (isSectionName(text)) {
			TheString::removeIniBracket(text);
			if (TheString::icmp(text,section)==0) {
				return i;		// section FOUND
			}	//endif
		}	//endif
		i++;
	}	//endwhile
	return -1;		// Section NOT found
}

//==============================================================================
//==============================================================================
int TheIniFile::findKey(const string& key,int lineNumber) {
	lineNumber++;	// since "lineNumber" points to [section], we need to skip 
					// to next line.
	string text;	// unaltered entire line being read.
	string keyName;		// current key name being read.
	int size=_currentIniFilePtr->_buffer.size();
	while (lineNumber<size ) {
		text=_currentIniFilePtr->_buffer.get(lineNumber);
		TheString::removeLeadingWhiteSpace(text);
		if (isSectionName(text)) {
			return -1;		// KEY Not found
		}
		if (isKeyName(text)) {
			extractKey(keyName,text);
			if (TheString::icmp(keyName,key)==0) {
				return lineNumber;		// key FOUND
			}
		}
		lineNumber++;
	}
	return -1;		// key NOT found
}

//==============================================================================
//==============================================================================
int TheIniFile::findSectionKey(const string& section,const string& key) {
	int i=findSection(section);	// i=section line #
	if (i<0) { 
		return -1;// not found
	}
	return findKey(key,i);	// return either line# or -1
}

//==============================================================================
//==============================================================================
bool TheIniFile::isSectionName(const string& section) {
	return (section[0]==SECTION_CHAR1); 
}

//==============================================================================
//==============================================================================
bool TheIniFile::isKeyName(const string& key) {
	return (!isSectionName(key) && !isComment(key));
}

//==============================================================================
//==============================================================================
bool TheIniFile::isComment(const string& line) {
	return (line[0]==COMMENT_CHAR1 || line[0]==COMMENT_CHAR2);
}

//==============================================================================
//==============================================================================
void TheIniFile::extractKey(string& key, const string& line){
	int position=line.find(KEY_DELIMITER_CHAR);	// ie look for '='
	if (position==string::npos) { key="";}	// error: no key exists.
	key.assign(line,0,position);
}
//==============================================================================
//==============================================================================
void TheIniFile::extractKeyValue( string& key, string& value, const string& line){
	int position=line.find(KEY_DELIMITER_CHAR);	// ie look for '='
	if (position==string::npos) {
		return;
		//return false;	// error: no key exists.
	}	
	key.assign(line,0,position);
	value.assign(line,position+1,string::npos);
	//return true;
}

//==============================================================================
//==============================================================================
void TheIniFile::extractKeyValue( string& key, string& value, int lineNumber ){
	 extractKeyValue(key,value,_currentIniFilePtr->_buffer.get(lineNumber));
}

//==============================================================================
//==============================================================================
void TheIniFile::extractKey(string& key,int lineNumber){ 
	extractKey(key, _currentIniFilePtr->_buffer.get(lineNumber));
}

//==============================================================================
//==============================================================================
bool TheIniFile::write(const string& section, const string& key, const string& value) {

	_currentIniFilePtr->setDirtyFlag(true);	// since we're writing, set dirty=1
	
	int lineNumber=findSection(section);
	string newSection=section;
	string keyValue=key+"="+value;
	// section doesn't exist.  Create section and key.
	if (lineNumber<0) {
		TheString::addIniBracket(newSection);
		_currentIniFilePtr->_buffer.push_back(newSection);
		_currentIniFilePtr->_buffer.push_back(keyValue);

	}

	// section exists.
	else {
		int keyLineNumber=findKey(key,lineNumber);
		// Key Doesn't exist.
		if (keyLineNumber<0) {
			lineNumber++;		// since lineNumber points to [section], 
								// go to next line.
			_currentIniFilePtr->_buffer.insert(lineNumber, keyValue);
			
		}
		// Key already exists
		else {
			_currentIniFilePtr->_buffer.set(keyLineNumber,keyValue);			
		}
	}

	//int status=WritePrivateProfileString(section.c_str(),key.c_str(),value.c_str(),_fileName.c_str());
	return 1;
}


//==============================================================================
//==============================================================================
bool TheIniFile::deleteKey(const string& section, const string& key) {
	int lineNumber=findSectionKey(section,key);
	// key doesn't exist.  Don't do anything.
	if (lineNumber<0) {
		return false;
	}

	_currentIniFilePtr->setDirtyFlag(true);
	_currentIniFilePtr->_buffer.erase(lineNumber);
	return true;
}

//==============================================================================
//==============================================================================
bool TheIniFile::deleteGroup(const string& section) {
	// 
	bool returnValue=false;		// returns true if any groups were deleted.
	int strLength=section.length();
	int i=0;
	string text;
	int size=0;
	while (i<_currentIniFilePtr->_buffer.size()) {
		text=_currentIniFilePtr->_buffer.get(i);
		// Filter only [Sections]
		if (isSectionName(text)) {
			TheString::removeIniBracket(text);	// remove "[ ]"
			// Find matching section names (that beginning)
			if (TheString::nicmp(text,section, strLength)==0) {
				// matches exactly (ie [section]==[section])
				if (TheString::icmp(text,section)==0) {
					size-=deleteSubGroup(i);
					returnValue=true;
					i--;	// 
				}
				// matches parent group only (ie [section]==[section/subsec/...])
				else if (text[strLength]=='/') {
					size-=deleteSubGroup(i);
					returnValue=true;
					i--;
				}
			}	//endif
		}	//endif
		i++;
	}	//endwhile
	return returnValue;		// Section NOT found
	
	return false;
}

int TheIniFile::deleteSubGroup(int lineNumber) {
	string text;
	bool stop=false;
	int lineDeleted=0;	// # of lines deleted.
	_currentIniFilePtr->setDirtyFlag(true);	// 
#ifdef _DEBUG	
	string _debugStr=_currentIniFilePtr->_buffer.get(lineNumber);
#endif
	_currentIniFilePtr->_buffer.erase(lineNumber);

	lineDeleted++;
	//lineNumber++;
	do {
		if (lineNumber >=_currentIniFilePtr->_buffer.getSize()) {
			stop=true;
		}
		else {
			text=_currentIniFilePtr->_buffer.get(lineNumber);
			if (isSectionName(text)) {
				stop=true;
			}
			else {
				_currentIniFilePtr->_buffer.erase(lineNumber);
				lineDeleted++;
				lineNumber--;	// we need to deduct 1 because when we 
								// erase(), the current "lineNumber" 
								// points to the next line already.
			}
		}
		lineNumber++;
	} while (!stop);
	
	return lineDeleted;

}
//==============================================================================
//==============================================================================
bool TheIniFile::flush() {
	if (!_currentIniFilePtr) {
		return false;
	}

	if (_currentIniFilePtr->getDirtyFlag()==false) {
		return false;		// nothing to flush.
	}
	string text="";
	int size=_currentIniFilePtr->_buffer.size();	// size of ini file
	int i=size-1;	

	while (i>=0) {
		text=_currentIniFilePtr->_buffer.get(i);
		if (text=="" || text==" ") {
			_currentIniFilePtr->_buffer.erase(i);
		}
		else {
			break;
		}
		i--;
	}
	

	// 1st, clean up the ini file by deleting any trailing empty space.
	
	// write the ini file to HD.
	_currentIniFilePtr->setDirtyFlag(false);	// no longer dirty.
	TheFile file;
	file.open(_fileName,"w");
	
	i=0;
	size=_currentIniFilePtr->_buffer.size();	// size of ini file
	while (i<size ) {
		file.writeLine(_currentIniFilePtr->_buffer.get(i));
		i++;
	}
	file.close();
	return true;
}

//==============================================================================
//
//==============================================================================
void TheIniFile::makeIniFile(string& iniPath, const string& fullPath) {
	char iniFilename[_MAX_PATH];
	strcpy(iniFilename, fullPath.c_str() );
	char *lastDot = strrchr( iniFilename, '.' );
	if ( lastDot ) {
		*lastDot = '\0';
	}
	strcat( iniFilename, ".INI" );
	iniPath = iniFilename;
}

//==============================================================================
//
//==============================================================================
void TheIniFile::dump() {
	string text="*** BEGIN ***\n";
	int size=_currentIniFilePtr->_buffer.size();
	int i=0;
	while (i<size ) {
		text=text+_currentIniFilePtr->_buffer.get(i)+"\n";
		i++;
	}
	text=text+"*** END ***";
	MessageBox(NULL,text.c_str(),_fileName.c_str(),MB_OK);
}


//
// OLD CODE
//
//
/*int IniFile::openSystem(const string file) {
	_filename= file;
	return 1;
}

int IniFile::openConfig() {
	string file=FileSystem::getAppFileName();
	string fileWithoutExtension=FileSystem::getFileNameWithOutExtension(file);
	string newFile=fileWithoutExtension+".ini";
	return open(newFile.c_str());
}
*/

/*
int  ConvertStrVar ( TCHAR* str , int size) {
	
	char* delim="%";

    SHORT  listPrcnt,               // List identifier
           nRet;                    // sentinal/called function return values
    STRING szTemp;     // Working variable
    BOOL   bNoPrcntAdd;             // Special case, don't add a '%'
    bNoPrcntAdd = TRUE;

  if ( (strchr(szStr, delim[0]) == NULL)  ) {
		// No expandable tokens...
		return 0;
	}

	token=new TCHAR[size+1];
	buffer=new TCHAR[size+1];
	buffer[0]='\0';
	TCHAR* tmpStrPtr=str;
	int i =0;
	while (tmpStrPtr[i]!=delim[0]) {
		buffer[i] = tmpStrPtr[i];
		i++;
	}
	token = strtok( str, delim );
	tmpStrPtr=new TCHAR[255];

	while ( NULL != token ) {
		StrToVar(token, tmpStrPtr);
		strcat (buffer, tmpStrPtr);
		lstStr.addElement( szBuf );
		token = strtok( NULL, str);	// strtok must use NUL for all calls except 1st
	}
	delete[] token;
	delete[] buffer;
	delete[] tmpStrPtr;


 
    str = "";     // We'll reconstruct this
    nRet = ListGetFirstString( listPrcnt, szTemp );
    if ( nRet = END_OF_LIST ) then
        // this should not happen since we checked to see if
        // there was at least 1 '%' in the string
    elseif ( nRet = -1 ) then
        // !!! error I dunno
    endif;
    if ( szTemp = "" ) then
        // First character in string was a '%', if we add it
        // we will have two % signs
        bNoPrcntAdd = TRUE;
    endif;

    while ( nRet = 0 )
        if ( !StrToVar( szTemp ) ) then
            // It didn't substite anything add the '%' back in
            //MessageBox("Setup error: doSubst cant substitute:"+szTemp,WARNING);
            if ( bNoPrcntAdd ) then
                str = str + szTemp;
                // Now we can start adding '%' signs again
                bNoPrcntAdd = FALSE;
            else
                str = str + "%" + szTemp;
            endif;
        else
            str = str + szTemp;
            // We just did a variable substitution, don't add the trailing '%'
            bNoPrcntAdd = TRUE;
        endif;
        nRet = ListGetNextString( listPrcnt, szTemp );
        if ( nRet = -1 ) then
            // !!! another error
        endif;
    endwhile;
    ListDestroy( listPrcnt );
    RemoveExcessBackSlash(str);
    return;
}
*/