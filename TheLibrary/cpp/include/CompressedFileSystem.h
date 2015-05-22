#ifndef _COMPRESSEDFILESYSTEM_H_
#define _COMPRESSEDFILESYSTEM_H_

#include "the.h"
//==============================================
//	TheCompressedFileSystem
//==============================================
class TheCompressedFileSystem : public TheFileSystem {
protected:
	TheFileSystem* _nativeFileSystem;	// Win/Mac/LinuxFileSystem
		// so it uses native fs for disk read/write op and also to
		// use its operation if the compressed file/dir is not available.
		//
public:
	TheCompressedFileSystem(const string& fileName,const string& mode,TheFileSystem* nativeFileSystem);
	virtual  int getFileList(StringVector& files){return 0;}
		// not implemented.

	virtual	 int makeShallowDirectory(const string& dir);	// make shallow dir 

	virtual  int copyFile(const string& src, const string& dest) {return 0;}

	virtual  int changeDirectory(const string& path){return 0;}	// "CD"
	virtual   string getCurrentDirectory(){return "";}	// get cwd
	
	virtual  setFileAttribute(const string& file, const string& attribute) {}
		// not fully implemented. Currently only supports (file, "+w");

	virtual int open(const string& fileName, const string& mode){return 0;}	// open file in its derived filesystem 
	virtual int close(){return 0;}
	virtual int read(){return 0;}
	virtual int seek(){return 0;}
	virtual int write(){return 0;}
	virtual int getSize(){return 0;}

};




#endif