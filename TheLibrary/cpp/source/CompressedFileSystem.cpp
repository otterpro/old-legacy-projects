#include "the.h"
TheCompressedFileSystem::TheCompressedFileSystem(const string& fileName,const string& mode,TheFileSystem* nativeFileSystem)
//	:_nativeFileSystem(nativeFileSystem) {
:TheFileSystem() {
}

int TheCompressedFileSystem::makeShallowDirectory(const string& dir) {
	// since it is not implemented yet, here is what it would look like
	// as an example.
	// 
	//	if (compression_mode==COMPRESS_IT) {
	//		compressDir(dir);
	//		nativeFileSystem->fprintf()...blah......;;;
	//	} else if (compression_mode==DONT_COMPRESS) {
	//		nativeFileSystem->makeShallowDirectory();
	//	}
	return 1;
}


