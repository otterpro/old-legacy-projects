#include "the.h"
#include <assert.h>
#include <sstream>
#include <io.h>	// chmod(), access(),
#include <stdio.h>	// remove()
// STATIC var init
TheFileSystem::Type TheFileSystem::_CurrentType=TheFileSystem::STANDARD;

TheFileSystem::TheFileSystem(const string& fileName) {
}

//________________________________________________________________________________
// 
//________________________________________________________________________________
int TheFileSystem::isPathSeparatorChar(const char c) {
	if (c==_pathSeparatorChar1 || c==_pathSeparatorChar2) {
		return 1;
	};	
	return 0;
}

//________________________________________________________________________________
// 
//________________________________________________________________________________
void TheFileSystem::stripFileExtension(string& ext, const string& path) {
	char pathStr[MAX_STRING];
	strcpy (pathStr, path.c_str());
	char* dotPtr = strrchr(pathStr,'.');
	*dotPtr=0;
	ext= pathStr;
}

//________________________________________________________________________________
// 
//________________________________________________________________________________
string TheFileSystem::replaceFileExtension(const string& fileName, 
										   const string& newExtension) {
	string drive, dir, file, ext;
	splitPath(fileName,drive,dir,file,ext);
	string newPath;
	makePath(newPath, drive,dir,file,newExtension);
	return newPath;
}

//________________________________________________________________________________
// 
//________________________________________________________________________________
void TheFileSystem::makePath(string& path, const string& drive,const string& dir,
							 const string& fileName, const string& ext) {
	char fullPath[MAX_STRING];

	string newDir=dir;		// dir: make sure to add trailing "/", ":", "\\", etc.
	string newFileName=fileName;	// file: make sure to remove any beginning "/", ":", "\\"

	addTrailingPathSeparator(newDir);
	removeInitialPathSeparator(newFileName);
	sprintf(fullPath,"%s%c%s%s",drive.c_str(),_rootSeparatorChar,newDir.c_str(),
		newFileName.c_str());
	if (ext!="") {
		if (ext[0]!='.') {
			strcat (fullPath,".");
		}
		strcat (fullPath,ext.c_str());
	}
	path=fullPath;
}

//________________________________________________________________________________
//  makeFullPathWithCurrentDirectory
//________________________________________________________________________________
void TheFileSystem::makeFullPathWithCurrentDirectory(string& path) {
	string oldFileName, 
		oldDrive,
		oldDir,
		oldExt,
		newPath;		// 

	splitPath((const string&)path,oldDrive, oldDir,oldFileName,oldExt);
	newPath=oldDir+oldFileName;

	string currentDir=getCurrentDirectory();
	string drive,dir,dummyStr,dummyExt;
	splitPath(currentDir,drive,dir,dummyStr,dummyExt);
	makePath(path, drive,dir,newPath,oldExt);

}


//________________________________________________________________________________
// makeFullPath
//________________________________________________________________________________
string TheFileSystem::makeFullPath( const char* dir, const char* fileName) {
	char fullPath[_MAX_PATH];
	strcpy(fullPath,dir);
	strcat(fullPath,_pathSeparator1.c_str());
	strcat(fullPath, fileName);
	string returnString = fullPath;
	return returnString;
}

//________________________________________________________________________________
//	makeFullPathWithAppDirectory
//________________________________________________________________________________
void TheFileSystem::makeFullPathWithAppDirectory(string& path) {
	string oldFileName, 
		oldDrive,
		oldDir,
		oldExt,
		newPath;		// 

	splitPath((const string&)path,oldDrive, oldDir,oldFileName,oldExt);

	string fullPath;
	TheOS::getAppFileName(fullPath);
	string drive,dir,dummyStr,dummyExt;
	splitPath(fullPath,drive,dir,dummyStr,dummyExt);

	newPath=oldDir+oldFileName;
	makePath(path, drive,dir,newPath,oldExt);
	
}

//________________________________________________________________________________
// makeFullPath
//________________________________________________________________________________
void TheFileSystem::makeFullPath(string& path) {
	if (isAbsolutePath(path)) { 
		return;
	}
	makeFullPathWithAppDirectory(path);
}

//________________________________________________________________________________
// isAbsolutePath
//	
//	very inefficient by using _splitPath()
//________________________________________________________________________________
bool TheFileSystem::isAbsolutePath(const string& path) {
	string drive,dir,fileName,ext;
	splitPath((const string&)path,drive, dir,fileName,ext);
	return (drive!="");	
}
//________________________________________________________________________________
//________________________________________________________________________________

void TheFileSystem::addTrailingPathSeparator(string& path) {
	if (path[path.length()-1]!=_pathSeparatorChar1 &&
		path[path.length()-1]!=_pathSeparatorChar2) {
		path=path+_pathSeparator1;
	}

}


//________________________________________________________________________________
//________________________________________________________________________________
void TheFileSystem::removeInitialPathSeparator(string& path) {
	if (path!="") {
		if (path[0]==_pathSeparatorChar1 ||
			path[0]==_pathSeparatorChar2) {
			path=path.substr(1,path.length()-1);
		}
	}
}

//________________________________________________________________________________
// 
//________________________________________________________________________________
void TheFileSystem::getDrive(string& drive, const string& path) {
	string dir,fileName,ext;
	splitPath(path,drive,dir,fileName,ext);
}

//________________________________________________________________________________
// 
//________________________________________________________________________________
void TheFileSystem::getFileNameWithOutExtension(string& fileName, 
												const string& path) {
	string drive,dir,ext;
	splitPath(path,drive,dir,fileName,ext);
}

//________________________________________________________________________________
// 
//________________________________________________________________________________
void TheFileSystem::getFileExtension(string& ext, const string& path) {
		string drive,dir,fileName;
	splitPath(path,drive,dir,fileName,ext);
}
//________________________________________________________________________________
// 
//________________________________________________________________________________
void TheFileSystem::getDir(string& dir, const string& path){
	string drive,fileName,ext;
	splitPath(path,drive,dir,fileName,ext);

}

//________________________________________________________________________________
// 
//________________________________________________________________________________
void TheFileSystem::getFileName(string& fileName, const string& path) {
		string drive,dir,ext;
	splitPath(path,drive,dir,fileName,ext);
	fileName=fileName+ext;
}
//________________________________________________________________________________
// 
//	makeDirectory
//
//	Traverses through dir structure and makes dir as needed.
//	It doesn't check if the dir exists first.   
//________________________________________________________________________________
int TheFileSystem::makeDirectory(const string& path) {
	// makes the directory and subdirectory if needed.
	string fullPath=path;
	char pathStr[MAX_STRING];
	strcpy (pathStr,path.c_str());
	// append "/" or "\\" at end before parsing because it won't get the last dir.
	// ex: "dir1\\dir2\\dir3" will only return "dir1\\dir2" because it thinks 
	// dir3 is a file.  rather, make it "dir1\\dir2\\dir3\\". 
	if (pathStr[strlen(pathStr)-1]!=_pathSeparatorChar1 || 
		pathStr[strlen(pathStr)-1]!=_pathSeparatorChar2) {
		fullPath = fullPath + _pathSeparator1;
	}
	string dir,currentPath;
	getDir(dir,fullPath.c_str());	// dir="dir1\\dir2\\dir3\\"
	getDrive(currentPath,fullPath.c_str());	// currentPath="D:\\"

	char delimiter[3];
	delimiter[0]=_pathSeparatorChar1;
	delimiter[1]=_pathSeparatorChar2;
	delimiter[2]=0;

	char str[MAX_STRING];
	strcpy (str,dir.c_str());
	char* strPtr=strtok(str,delimiter);
	while (strPtr) {
		currentPath=currentPath + _pathSeparator1 + strPtr;
		makeShallowDirectory(currentPath);
		
		strPtr=strtok(NULL,delimiter);
	}
	return 1;
}

#ifdef WIN32
#define _THE_FILE_SYSTEM TheWinFileSystem
#endif
#ifdef __MWERKS__
#define _THE_FILE_SYSTEM TheMacFileSystem
#endif
#ifdef _LINUX	//??? is this macro correct?
#define _THE_FILE_SYSTEM TheLinuxFileSystem
#endif

//________________________________________________________________________________
//
//________________________________________________________________________________
TheFileSystem* TheFileSystem::create(const Type type) {
	if (type==STANDARD) {
		TheFileSystem* fileSystem = new _THE_FILE_SYSTEM();
		//fileSystem->open(fileName, mode);
		return fileSystem;
	}
	else if (type==COMPRESSED) {
		// TODO: uncomment below after finishing compressedFileSystem 
		// return new TheCompressedFileSystem(_FileSystem);
	}
}
// GLOBAL INSTANCE
TheFileSystem* _FileSystem=new _THE_FILE_SYSTEM();

//________________________________________________________________________________
//
//________________________________________________________________________________
int TheFileSystem::open(const string& path, const string& mode) {
	int ioMode=0;
	// "R" for READ
	if (mode.find("r") !=string::npos) {
		ioMode=ioMode|ios_base::in;
	}
	// "W" or "A" but cannot have both
	if (mode.find("w") !=string::npos) {
		ioMode=ioMode|ios_base::out|ios_base::trunc;
	}
	else if (mode.find("a") !=string::npos) {
		ioMode=ioMode|ios_base::out|ios_base::app;
	}

	_file.open(path.c_str(), ioMode);
	return _file.good();
}
//________________________________________________________________________________
//
//________________________________________________________________________________
int TheFileSystem::close() {
	if (_file.is_open()) {
		_file.close();
		return 1;		// file closed.
	}
	return 0;	// no file needed to close
}
//________________________________________________________________________________
//
//________________________________________________________________________________
void TheFileSystem::readStr(string& str) {
	//assert(_file);
	_file>>str;
	//return str;
}
//________________________________________________________________________________
//
//________________________________________________________________________________
int TheFileSystem::readInt() {
	string str;
	readStr(str);
	return atoi(str.c_str());
}

//________________________________________________________________________________
//
//________________________________________________________________________________
void TheFileSystem::readLine(string& str) {
	char buffer[MAX_STRING];
	_file.getline(buffer,MAX_STRING);
	//string str=buffer;
	//return str;
	str=buffer;
}

//________________________________________________________________________________
//
//________________________________________________________________________________
int TheFileSystem::writeLine(const string& buffer) {
	_file<<buffer<<'\n';
	return 1;
}
//________________________________________________________________________________
//
//________________________________________________________________________________
int TheFileSystem::write(string& word) {
	_file<<word;
	return 1;
}
//________________________________________________________________________________
// 
//________________________________________________________________________________
int TheFileSystem::write(int number){
	std::ostringstream str;
	str << number;
	return write(str.str());
}

//________________________________________________________________________________
//________________________________________________________________________________
int TheFileSystem::remove(const string& fileName) {
	return remove(fileName.c_str());
}

//________________________________________________________________________________
//
//________________________________________________________________________________
int TheFileSystem::chmod(const string& fileName,const string& permission) { 
	// TODO: not implemented yet
	// Apple CodeWarrior 3 doesn't have chmod in io.h

	//string permissionString=permission;
	//TheString::toLower(permissionString);
	//if (permissionString.find_first_of("r")!=string::npos) {
	//}
	return 0;
}
//________________________________________________________________________________
//
//________________________________________________________________________________
int TheFileSystem::rename(const string& oldFileName, const string& newFileName) {
	return rename(oldFileName.c_str(),newFileName.c_str());		
}


//________________________________________________________________________________
//
//________________________________________________________________________________
int TheFileSystem::fileExist(const string& file) {
	//return ((access(file.c_str(), F_OK)+1)?true:false);
		// access() might be WIN only
		// Apple CodeWarrior 3 doesn't have access() in io.h

	// OLD code that didn't use <io.h> access()
	FILE *fp;
    fp = fopen( file.c_str(), "r" );
    if( fp != NULL ) {
	fclose( fp );
	  return 1;
    }
	return 0;
}

//int TheFileSystem::canRead(const string& fileName) {}


