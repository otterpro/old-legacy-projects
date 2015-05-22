#include "the.h"
//=====================================================
//
//=====================================================
int TheWinFileSystem::getFileList(StringVector& files) {
	return 0;
}

/*
string TheWinFileSystem::makePath(const string drive,const  string dir, const string fileName, const string ext) {
	char fullPathStr[MAX_STRING];
	_makepath(fullPathStr,drive.c_str(),dir.c_str(),fileName.c_str(),ext.c_str());
	string fullPath=fullPathStr;
	return fullPath;
}
*/

//=====================================================
//=====================================================
TheWinFileSystem::TheWinFileSystem():TheFileSystem() {
	_pathSeparatorChar1='\\';	// path separator #1
	_pathSeparatorChar2='/';	// #2
	_pathSeparator1="\\";	// same as above but in string
	_pathSeparator2="/";	
	_rootSeparatorChar=':';	// Letter that separates root 
	// in Linux = '/'.  Mac=':'  PC=':'.  Ie in Mac, "VolName:blah:bla"
	// in PC, "C:\blah\blah".  In Linux, "/usr/blah".   
}
/*

	pathSeparatorChar1='\\';
	pathSeparatorChar2='/';
	pathSeparator1=pathSeparatorChar1;
	pathSeparator2=pathSeparatorChar2;

}
*/


//=====================================================
//=====================================================
/*string TheWinFileSystem::makeFullPath(const string& file) {

	char fullPathStr[_MAX_PATH];
	_fullpath(fullPathStr,file.c_str(),_MAX_PATH);
	string fullPath = fullPathStr;
	return fullPath;
}
*/

//=====================================================
//=====================================================
/*string TheWinFileSystem::findFirstDir(const char* dir, const char* nameToMatch) {
	WIN32_FIND_DATA findFileData;

	char pathToSearch[_MAX_PATH];
	sprintf(pathToSearch,"%s%c%s",dir,pathSeparatorChar1,nameToMatch);

	_findDirHandle=FindFirstFile(pathToSearch,&findFileData);
	if (_findDirHandle==INVALID_HANDLE_VALUE) {
		return "";
	}
	if (findFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY && 
		stricmp(findFileData.cFileName,".")!=0 && 
		stricmp(findFileData.cFileName,"..")!=0) {
			string returnString =findFileData.cFileName;
			return returnString;
	} else {
		// keep looking for dir and make sure it matches our given "nameToMatch" param.
		return findNextDir();
	}
}

//=====================================================
//=====================================================
string TheWinFileSystem::findNextDir() {
	WIN32_FIND_DATA findFileData;
	int	status=FindNextFile(_findDirHandle,&findFileData);

		if (!status) {
				// ERROR: no matching dir was found.
			FindClose( _findDirHandle );
			return "";
		}
		else if (findFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY && 
				stricmp(findFileData.cFileName,".")!=0 && 
				stricmp(findFileData.cFileName,"..")!=0) {
					string returnString =findFileData.cFileName;
					return returnString;
		} else {
			// keep looking.
			return findNextDir();	// recursive call
		}
}

//=====================================================
// warning: if filename 
//=====================================================
string TheWinFileSystem::findFirstFile(const char* dir, const char* nameToMatch) {
	WIN32_FIND_DATA findFileData;
	bool isFile=false;
	
	char pathToSearch[_MAX_PATH];
	sprintf(pathToSearch,"%s%c%s",dir,pathSeparatorChar1,nameToMatch);

	_findFileHandle=FindFirstFile(pathToSearch,&findFileData);
	if (_findFileHandle==INVALID_HANDLE_VALUE) {
		return "";
	}

#ifdef _DEBUG
	DWORD a=findFileData.dwFileAttributes;
	DWORD b=FILE_ATTRIBUTE_DIRECTORY;
#endif
	// if it is DIR and not a file, keep looking.
	if (findFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY ) {
		return findNextFile();
	}

	string returnString =findFileData.cFileName;
	return returnString;
}




string TheWinFileSystem::findNextFile() {
	WIN32_FIND_DATA findFileData;

	int status=FindNextFile(_findFileHandle,&findFileData);

#ifdef _DEBUG
	DWORD a=findFileData.dwFileAttributes;
	DWORD b=FILE_ATTRIBUTE_DIRECTORY;
#endif

	if (!status) {
		// Error occured
		FindClose( _findFileHandle );
		return "";
	}

	
	else if (findFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
		// It is a DIR and not FILE
		return findNextFile();	// recursive call
	} 
	else {
		// FOUND IT
		string returnString =findFileData.cFileName;
		return returnString;
	}
}
*/
//=====================================================
//
//=====================================================
int TheWinFileSystem::makeShallowDirectory(const string& dir) {
	return CreateDirectory(dir.c_str(),NULL);
}
//=====================================================
//
//=====================================================
int TheWinFileSystem::copyFile(const string& src, const string& dest){
	return CopyFile(src.c_str(),dest.c_str(),FALSE);	// FALSE = overwrite. TRUE= don't overwrite.
}
//=====================================================
//
//=====================================================
string TheWinFileSystem::getCurrentDirectory() {
	// WIN
	char currentDir[MAX_STRING];
	currentDir[0]=0;
	GetCurrentDirectory(MAX_STRING,currentDir);
	string returnString=currentDir;
	addTrailingPathSeparator(returnString);
	return returnString;
}
//=====================================================
//
//=====================================================
int TheWinFileSystem::changeDirectory(const string& path) {
	return SetCurrentDirectory(path.c_str());
}

//=====================================================
//
//=====================================================
int TheWinFileSystem::setFileAttribute(const string& file, const string& attribute) {
	if (attribute=="+w") {
		SetFileAttributes(file.c_str(),FILE_ATTRIBUTE_NORMAL);
	}
}
