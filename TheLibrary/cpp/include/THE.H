//
#ifndef _THE_H_
#define _THE_H_

// WIN CE - POCKET PC 
#ifdef _WIN32_WCE	
	#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers
	#include <windows.h>	//non-MFC
	#include <afx.h>	// MFC, using CString (doesn't use CObject), TRACE, ASSERT
	#include <commctrl.h>
	#include <aygshell.h>
	#include <sipapi.h>		// SIP - input panel

// WIN 32 - 9x/NT/2k/XP
#elif _WIN32
	#include <windows.h>	//non-MFC
#else

	#define DONT_USE_STREAM 1
	// if defined, it uses STDIO instead of FSTREAM.h
	//#ifndef DONT_USE_STREAM

	#define INT2BOOL()

	#include <string>
	using namespace std;

	//#include <fstream.h>
	#include <fstream>
	//#endif
	#include <string.h>

	#define COMMENT_CHAR1 '#'
	#define COMMENT_CHAR2 ';'



		// an integer can go upto 32 digit in our lib.
	//==============================================================================
	//	OS related static/global functions
	//==============================================================================
	class TheOS {
	public:
		TheOS();
		~TheOS();
		static void splitPath(const string& path, string& drive, string& dir, 
							string& fileName, string& ext);
							// splits a full path into a small components

		static int launch (const string& file, const string& arg="",
							const string& workingDir="", int wait=false);

		static void getAppFileName(string& fileName);	
									// get the running app's fullPath
									// for just the path, use 
									//		TheFileSystem::getAppPath()
		static void getSystemPath(string& path);
									// get system directory
									// Win: "C:\windows\" (or %windir%)
									// Mac: "System Folder:Preferences:"
									// Linux: "\etc\"
	};

	// TODO:
	class TheWinOS: public TheOS {
	};

	class TheMacOS: public TheOS {
	};

	class TheLinuxOS: public TheOS {
	};

	extern TheOS* _OS;



	#include "misc.h"

	#include "filesystem.h"


	#include "stringvector.h"

	#include "mmedia.h"

	#ifdef _DEBUG
		#define DebugStr(x){MessageBox(NULL,x,"Debug");}
		#define DebugInt(x){char dbgMsg[MAX_STRING];sprintf(dbgMsg,"%d",x);DebugStr(dbgMsg);}
	#else
		#define DebugStr(x){}
		#define DebugInt(x){}
	#endif


	#include "configfile.h"
	#include "inifile.h"
	#include "reg.H"
	#include "filesystem.h"  
	#include "winfilesystem.h"
	#include "compressedfilesystem.h"
#endif // 





//==============================================================================
//==============================================================================
/*class LogFile {
public:
#ifdef DONT_USE_STREAM
	FILE* _file;
#else
	ofstream _file;
#endif
	string _filename;
	LogFile();
	virtual ~LogFile() { close(); }
	int open(const string file, int mode=0);
		// mode 0 = normal, mode 1 = append
	int writeln(const string text);
	void close();
};

//==============================================================================
//==============================================================================
class DebugLogFile : public LogFile {
#ifndef _DEBUG
	int open(const char* file, int mode=0) {}
	int write(const char* text){}
	void close(){}
#endif
};
*/



#ifdef USE_APP_IS_NOT_FINISHED
//==============================================================================
//	currently not used
//==============================================================================
class TheApp{
public:
	TheConfigFile* _configFile;	// app.INI should be accessed from here.
								// don't create TheConfigFile directly unless
								// it is not app.INI.  
	//string _fullPath;	// full path of this executable (path+filename)
	
	TheApp();	// init _FileSystem and ::_configFile
	~TheApp();

	// alias func
	static void getAppFileName(string& fileName) {	
								// get the running app's filename.
								TheOS::getAppFileName(fileName);
								}
	int showLastError();		// show any error message, if any

};

extern TheApp _App;	// _App is global.  Don't forget to call init() in WinMain() or main().
#endif

#endif // _WINTOOL_H_
