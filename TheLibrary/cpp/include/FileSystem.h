#ifndef _FILESYSTEM_H_
#define _FILESYSTEM_H_

#include "the.h"
//#include <assert.h>

// Declaration
class TheFile;


//==============================================================================
//	FileSystem
//==============================================================================
class TheFileSystem {
protected:
	// const
	 char _pathSeparatorChar1;	// path separator #1
	 char _pathSeparatorChar2;	// #2
	 string _pathSeparator1;	// same as above but in string
	 string _pathSeparator2;	
	 char _rootSeparatorChar;	// 

	 fstream _file;			// This File Stream 

	 string _fileName;		// This FileName

	//static char pathSeparators[3];	// contains pathSep1 , pathSep2 as delimiters

	// WIN ONLY (won't be used after LS() is done, replaced with StringVector())
	//HANDLE _findFileHandle;	// used by findFirstFile(), findNextFile();
	//HANDLE _findDirHandle;	// used by findFirstDir(), findNextDir()

public:
	TheFileSystem(const string& fileName="");
	int isPathSeparatorChar(const char c);	// returns 1 if "c" is ":" or "//", 
					//	depending on setup could be inlined.
	int xCopy(const string& src, const string& dest, int mode=0);	
							// perform xcopy
							// currently, param "mode" is ignored.

	virtual string replaceFileExtension(const string& fileName, 
										const string& newExtension);
							// replace..("helpme.doc",".txt"); 
							// makes it "helpme.txt"
							// Not implemented yet.
	void stripFileExtension(string& fileName, const string& path);	
							// given a path, remove any extension of file/dir.
							// ex: "helpme.doc" -> "helpme". Works on DIR as well.
	string getNextFile(const string& path="");
	 // TODO:
				// if param "path" is defined, it first does "ls" 
				// and retrieves 1st file/dir.
				// If "path" is not defined, it retrieves 2nd or next file/dir.
				// If there are no more files, it returns "".
				// Directory ends in "PathSeparatorChar1" and 
				// filename doesn't end in it.
	bool isAbsolutePath(const string& path);	//
								// returns true if given
								// path is full (absolute) path 
	void makePath(string& path, const string& drive,
						const string& dir, const string& fileName, 
						const string& ext="");
					// assembles the component of path and returns full path
	void makeFullPathWithCurrentDirectory(string& path);
	void makeFullPathWithAppDirectory(string& path);
					// given a path, make it a full path using app's path.
					// if the path is already a full path, then it will extract the 
					// filename and use the app directory instead of current path.
					// ex: "blah.txt" => "d:\me\blah.txt" (exe is in d:\me\)
					// ex2: "e:\tool\blah.txt" => "d:\me\blah.txt" (exe is in d:\me\)
	  string makeFullPath(const char* dir, const char* fileName);	
							// make a fullPath using param1 (dir)
						// REDO:

	void makeFullPath(string& path);
					// if the path is already a full path (absolute path), it does nothing.
					// if "path" is only a filename or directory and filename, it uses the
					// AppDirectory() ie appends the path to AppDirectory()
					// ex: ("temp\blah.txt") => "C:\mydir\temp\blah.txt", 
					// (executable is in c:\mydir\) 

	void addTrailingPathSeparator(string& path);
					// adds trailing path separator (ie ":" or "\\" or "/" to path) 
					// if the path lacks it.
		
	void removeInitialPathSeparator(string& path);
					// if the path begins with path separator, remove it.
					// ex: (linux) "/blah/blah2"-> "blah/blah2".  
					// "blah2/blah3"->"blah2/blah3" (no change)

	void getDrive(string& drive, const string& path);	
				// get volume name (mac) or drive letter (pc) or root dir (linux)
				// Win ex: (C:\sub\file) => "C" 
				// Linux:  (\etc\blah) => "\etc"
				// Mac:		(Vol1:sub:file)	=>"Vol1"


	void getFileName(string& fileName, const string& path);

	void getFileNameWithOutExtension(string& fileName, const string& path);

	void getFileExtension(string& ext, const string& path);

	void getDir(string& dir, const string& path);	
						// get dir path given a path ("C:/dir/blah.exe"->"dir/")
	
	int makeDirectory(const string& path);	
							// makes the directory and subdirectory if needed.
							// ex: if there is /dir1/, makeDir("/dir/sub1/sub2"); 
							// will create 2 subfolders-- "sub1" and "sub2".
							// If Dir already exists, it does nothing.
	// not implemented
	const string& getFirstFile(const string& path);
	const string& getNextFile();
	const string& getFirstDir();
	const string& getNextDir();

	// Alias func (TheOS)
	void splitPath(const string& path, string& drive, string& dir, 
							string& fileName, string& ext) {
							TheOS::splitPath(path,drive,dir,fileName,ext);}
							// splits a full path into a small components

	int launch (const string& file, const string& arg="", 
						const string& workingDir="", int wait=false) {
						return TheOS::launch(file,arg,workingDir,wait);}

	void getAppPath(string& path) {
							// get this app's path only 
							string fullPath, drive,dir,fileName,ext;
							TheOS::getAppFileName(fullPath);
							splitPath(fullPath,drive,dir,fileName,ext);
							makePath(path,drive,dir,"","");
	}								

	// either uses ANSI or custom implementation for Compressed File
	virtual int remove(const string& fileName);
	//virtual int remove(const TheFile& file);


	virtual int chmod(const string& fileName,const string& permission);	// not implemented yet
	//virtual int chmod(const TheFile& file,const string& permission);

	
	virtual int rename(const string& oldFileName, const string& newFileName);

	// TODO:
	//virtual int canRead(const string& fileName);
	//virtual int canWrite(const string& fileName);
	


	//--------------------------------------
	// defined at platform level
	//--------------------------------------


	// not implemented
	//virtual  int getFileList(StringVector& files);

	// warning: find___() can be dangerous bc more than 1 concurrent find___() will result in
	//	corrupted state.  Could implement stack instead to keep track or use "ls()" instead.
	// USE getNextFile() instead.
	// string findFirstDir(const char* dir, const char* nameToMatch="*.*");

	// string findNextDir();	// find next dir

	//string findFirstFile(const char* dir, const char* nameToMatch="*.*");

	//string findNextFile();	// find next dir



	virtual	 int makeShallowDirectory(const string& dir)=0;
								// make shallow dir, ie will not create nested
								// directory.   if the nested dir doesn't exist.
								//Used by makeDirectory().
	virtual  int copyFile(const string& src, const string& dest)=0;
								// make shallow copy.  Used by XCopyFile().

	virtual  int changeDirectory(const string& path)=0;	
								// "CD"
	virtual   string getCurrentDirectory()=0;	
								// get cwd
	virtual setFileAttribute(const string& file, const string& attribute) {}
								// not fully implemented. 
								// Currently only supports (file, "+w");
	// FILE FUNCTIONS, depending on FileSystem (either using Std lib, WX, or 
	//					Compressed File System)
	virtual int fileExist(const string& file);	// returns 1 if file exists.
	int exist(const string& file) { return fileExist(file);}	// alias of fileExist.

	// File System Types
	enum Type{ STANDARD, COMPRESSED, ZIP, LHA};	// ZIP, LHA not supported.  Just used as an example.

protected:
	static Type _CurrentType;	// STANDARD, COMPRESSED, etc.  By default, STANDARD

public:
// GENERAL File I/O
	static TheFileSystem* create(const Type type=STANDARD);
										// Creates the FileSystem/File dynamically.

	static Type getCurrentType() { return _CurrentType;}

	// USING Standard IO, override with Compressed IO, etc.

	virtual int open(const string& fileName, const string& mode);	
									// open file in its derived filesystem 
	virtual int close();			// close file

	virtual void readLine(string& str);	// read 1 line of text file
	virtual void readStr(string& str);	// read 1 word of text file
	int readInt();						// read 1 integer of text file
	
	virtual int writeLine(const string& buffer);	
	virtual int write(string& word);
	int write(int number);

	virtual int eof() 
								{ return _file.eof();}
	virtual void flush()		{_file.flush();}
};

// Global Instance of TheFileSystem
//	Make sure to init it during TheApp ctor()
extern TheFileSystem* _FileSystem;


//==============================================================================
//	TheFile
//	NOTE:  TheFile:: is only an alias to TheFileSystem::
//==============================================================================
/*class TheFile {
public:
protected:
	TheFileSystem* _fileSystem;
	string _fileName;

};
*/
// OLD CODE
//#ifdef OLD
class TheFile /*: public TheFileSystem */{
private:
	TheFileSystem* _fileSystem;
	string _fileName;
	void init() {
		bool _init=false;
		if (!_init) {
			_fileSystem=TheFileSystem::create(TheFileSystem::getCurrentType());
			_init=true;
		}
	} 
	
	enum {NONE=0, READ=1, WRITE=2, APPEND=4 };	// 2 different mode
	int _mode;		// READ | WRITE | APPEND
public:
	TheFile() : _fileSystem(0),_mode(0) {}
	~TheFile() {
		if (_fileSystem) {
			close();	// close any open file
		}
	}
	int open(const string& fileName, const string& mode="r") {
		init();
		string newMode=mode;
		TheString::toLower(newMode);

		int i=newMode.find("r");
		if (i!=string::npos) {
			_mode=_mode|READ;
		}
		i=newMode.find("w");
		if (i!=string::npos) {
			_mode=_mode|WRITE;
		}
		return _fileSystem->open(fileName,mode);
	}
	int close() {
		return _fileSystem->close();
	}
	void readStr(string& str) {
		_fileSystem->readStr(str);
	}
	void readLine(string& str) {
		_fileSystem->readLine(str);
	}
	int readInt() {
		return _fileSystem->readInt();
	}
	int writeLine(string& buffer) {
		assert(_mode&WRITE);	// make sure it can write.
		return _fileSystem->writeLine(buffer);
	}
	int write(string& word) {
		return _fileSystem->write(word);
	}

	int write(int number) {
		return _fileSystem->write(number);
	}
	int eof() { return _fileSystem->eof();}
	void flush() {
		_fileSystem->flush();
	}
};
//#endif

#endif