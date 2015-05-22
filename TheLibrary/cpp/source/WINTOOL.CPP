//#include <fstream.h>
#include <stdio.h>

#include "the.h"
#ifdef WIN32
#include <windows.h>	// Used by WinFileSystem::

/*
LogFile::LogFile() {
#ifdef DONT_USE_STREAM
	_file = 0;
#endif

}
//==============================================================================
//
//==============================================================================
int LogFile::open(const string& file, int mode) {
	// mode 0 = normal, mode 1 = append
	string path=file;
	_FileSystem->makeFullPath(path);	

	close();	// first close if needed
	if (mode==0) { 
#ifdef DONT_USE_STREAM
	_file = fopen(path.c_str(), "w");
#else
		_file.open(path, out);
#endif
	} 
	else {
#ifdef DONT_USE_STREAM
	_file = fopen(path, "a");
#else
		_file.open(path, out|app);
#endif
	}
	return 1;
}
int LogFile::writeln(const string text){
#ifdef DONT_USE_STREAM
	fprintf(_file,"%s\n",text.c_str());
#else
	_file<<text.c_str() <<endl;
#endif
	return 1;
}

void LogFile::close() {
#ifdef DONT_USE_STREAM
	if (_file) {
		fclose(_file);
	}
#else
	_file.close();
#endif
}

*/


//==============================================================================
//
//==============================================================================
void TheOS::splitPath(const string& path, string& drive, string& dir, 
							 string& fileName, string& ext) {
	
   char driveStr[_MAX_DRIVE];
   char dirStr[_MAX_DIR];
   char fileNameStr[_MAX_FNAME];
   char extStr[_MAX_EXT];

   _splitpath( path.c_str(), driveStr, dirStr, fileNameStr, extStr );
	drive=driveStr[0];
	fileName=fileNameStr;
	dir=dirStr;
	ext=extStr;
	//return 1;
}




//==============================================================================
//
//==============================================================================
void TheOS::getAppFileName(string& fileName) {
	char appPath[MAX_STRING];
	GetModuleFileName(NULL,appPath,MAX_STRING);
	fileName=appPath;

}

//==============================================================================
//
//==============================================================================
int TheOS::launch(const string& file, const string& arg, const string& workingDir, int wait) {
#ifdef WIN32
	if (file=="") return 0;

	char *tmpWorkingDir=0;
	int retval;	// 
	SHELLEXECUTEINFO shellExecInfo;

	// Prepare /Init
	shellExecInfo.cbSize = sizeof(shellExecInfo);
	
	//shellExecInfo.fMask = SEE_MASK_FLAG_NO_UI | SEE_MASK_NOCLOSEPROCESS;    // No errors
	shellExecInfo.fMask = SEE_MASK_FLAG_NO_UI ;    // No errors
			// change this fMask to appropriate behavior
	shellExecInfo.hwnd = NULL;  // System window gets errors
	shellExecInfo.lpVerb = NULL;    // Default to open
	shellExecInfo.nShow = SW_SHOWNORMAL;

	shellExecInfo.lpFile = file.c_str();
	shellExecInfo.lpParameters = arg.c_str();

	// Select working directory
	// case1: working dir is defined.
	if (workingDir!="") {
		shellExecInfo.lpDirectory = workingDir.c_str();
	} 
	else {
		tmpWorkingDir = new char[(file.size())+1];
		wsprintf( tmpWorkingDir, "%s", file );
		char *endPointer = strrchr(tmpWorkingDir, '\\' );
		if ( endPointer != NULL )
		{	// case2: use the path of the file/exec as working directory 
			*endPointer = '\0';
			shellExecInfo.lpDirectory = tmpWorkingDir;
		}
		else
		{	//case3:  Error occured.  use the current directory as working directory
			shellExecInfo.lpDirectory = NULL;	
		}
		
	}

	//printDbgMsg( shellExecInfo.lpFile, shellExecInfo.lpParameters );

	// Launch!
	retval = ShellExecuteEx( &shellExecInfo );

	if (tmpWorkingDir) {
			delete[] tmpWorkingDir;
	}

	//showShellExecErr( &shellExecInfo );
	//ghInstApp = shellExecInfo.hInstApp;	// Returns the hInst of the newly started app
	//ghProcess = shellExecInfo.hProcess;

	// TODO:  analyze the  shellExecInfo.hInstApp to see if there is any error!
	//
	if (retval >=32) { 
		return 0;	// SUCCESS
	}
	else {
		return -1;	// error. couldn't launch.
	}
#endif
}

//==============================================================================
//
//==============================================================================
void TheOS::getSystemPath(string& path) {
	char winDir[MAX_STRING];
	GetWindowsDirectory(winDir,MAX_STRING);
	// append "\\" if needed
	if (winDir[strlen(winDir)-1]!='\\') {
		strcat(winDir,"\\");
	}
	path=winDir;
};
	
//==============================================================================
//
//==============================================================================
int TheApp::showLastError() {
	// win only
	DWORD errorCode=GetLastError();
	if (!errorCode) { return 0;}
	
	char text[MAX_STRING];
	FormatMessage( FORMAT_MESSAGE_FROM_SYSTEM,
			NULL, GetLastError(), NULL, text, MAX_STRING, NULL );
	MessageBox( NULL, text, "GetLastError", MB_OKCANCEL );
	
	return 1;
}

//==============================================================================
//
//==============================================================================
TheApp::TheApp() {
#ifdef WIN32
	//Win Only

	// get _fullPath of app
	//char appPath[MAX_STRING];
	//GetModuleFileName(NULL,appPath,MAX_STRING);
	//_fullPath=appPath;

	// init global _FileSystem
	_FileSystem = new TheWinFileSystem; 

	// init _configFile

#endif
}

//==============================================================================
//
//==============================================================================
TheApp::~TheApp() {
	delete _FileSystem;
	//delete _configFile;
#ifdef WIN32

#endif
}

//App _App;	// _App is GLOBAL
#endif	// WIN32