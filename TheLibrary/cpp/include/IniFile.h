#ifndef _THEINIFILE_H
#define _THEINIFILE_H

#include <string>
using namespace std;
#include "the.h"

// INI Section Char
#define SECTION_CHAR1 '['
#define SECTION_CHAR2 ']'
#define KEY_DELIMITER_CHAR '='

//==============================================================================
//	TheIniFileBuffer (used by TheIniFile)
//	
//	holds the entire content of 1 unique ini file in the memory.	
//==============================================================================
class TheIniFileBuffer {
protected:
	string _fileName;		// fully qualified ini file name 
	bool _dirtyFlag;			// 
public:
	string& getFileName() { return _fileName;}
	StringVector _buffer;	// actual text
	void setFileName(const string& fileName) {_fileName=fileName;}
	bool getDirtyFlag() { return _dirtyFlag;}
	void setDirtyFlag(int flag) {_dirtyFlag=(bool)flag;}
	TheIniFileBuffer(): _fileName(""),_dirtyFlag(false) {}
	~TheIniFileBuffer() { clear();}
	void clear() { _buffer.clear(); }
};

//==============================================================================
//	TheIniFile
//	
//	Ini File manager which keeps track of all the open ini file in the memory.
//
//==============================================================================
class TheIniFile : public TheConfigFile {
	
public:
	TheIniFile();
	~TheIniFile();
	virtual bool open (const string& filename);		// opens ini file.  
						//  filename: if not using fullpath, it uses the current 
						//				app's path.

	virtual bool openConfig();	//	opens app's ini file ie
						//	"myApp.exe" -> opens "myApp.INI"

	virtual bool openSystem(const string& filename);	// Looks for this ini file in
						// Win: "C:\windows\" (or windows dir)
						// Mac: "Preferences Folder"
						// Linux: "\etc\"
	
	virtual void readStr(string& str, const string& group, const string& key, 
							const string& defaultString="");
	virtual bool write(const string& group, const string& key,
						const string& value/*,int index*/);
	virtual bool deleteKey(const string&  group, const string& key/*, int index*/);
							// delete entire key and all its values
	virtual bool deleteGroup(const string& group);
							// delete entire section, key.  
	virtual bool flush();
	virtual bool close();

	// misc string functions.
	static bool isSectionName(const string& section);
							// is it "[xxxxx]"?
	static bool isKeyName(const string& key);
							// true if it is in the form "_key=..."
	static bool isComment(const string& line);
							// true if it starts with ";", "#" 
	static void makeIniFile(string& iniPath, const string& fullPath);
				// change file extension to ini (ie "blah.txt"-> "blah.ini")

	// debug function
	void dump();		// show the current content of INI file to screen.

protected:
	TheFile _file;	// physical file 
	//StringVector _buffer;
	static vector <TheIniFileBuffer*> _iniDB;	
							// all the _iniDB.  Static/Global
	TheIniFileBuffer* _currentIniFilePtr;	// current INI FILE
	bool clear();			// clear this ini from memory
	void clearAll();		// clear all the _iniDB from memory.
	int findSection(const string& section);	// searches for [section] 
								// & returns line # or
								// -1 if not found.
	int findKey(const string& key, int lineNumber);
								// finds the key starting the lineNumber 
								// until next sectin is reached.
								// Returns line# or -1  if not found
	int findSectionKey(const string& section, const string& key);	
								// searches for [section],key=
								// and returns line# or 
								// -1 if not found.
	void extractKey(string& key, int lineNumber);
	 void extractKey(string& key, const string& line);
	
	 void extractKeyValue(string& key, string& value, const string& line);
	 void extractKeyValue(string& key, string& value,int lineNumber );

	int deleteSubGroup(int lineNumber);	// deletes all entries starting
						// this line and doesn't stop deleting until it
						// reaches the next [section].
						// It will delete comments as well.  
						// Returns # of lines that were deleted.
};

//==============================================================================
//	Creates Config File Instance and 
//		Manages them
//==============================================================================
/*class TheConfigFileFactory {
private:
	static vector<TheConfigFile*>_configDB;
public:
	static TheConfigFile* open(const string filename);


};
*/
#endif