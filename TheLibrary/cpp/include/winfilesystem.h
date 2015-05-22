#ifndef _WINFILESYSTEM_H_
#define _WINFILESYSTEM_H_

#include "the.h"
//==============================================
//	WinFileSystem
//==============================================
class TheWinFileSystem : public TheFileSystem {
public:
	TheWinFileSystem();
	virtual  int getFileList(StringVector& files);
		// not implemented.

	virtual	 int makeShallowDirectory(const string& dir);	// make shallow dir 

	virtual  int copyFile(const string& src, const string& dest);

	virtual  int changeDirectory(const string& path);	// "CD"
	virtual   string getCurrentDirectory();	// get cwd
	
	virtual  setFileAttribute(const string& file, const string& attribute);
		// not fully implemented. Currently only supports (file, "+w");

};




#endif