global class:
TheFileSystem* _FileSystem 
	(instantiation of TheWin/Mac/LinuxFileSystem*)
TheApp* _App
	(instantiatin of TheApp, see file Winapp.cpp,macapp.cpp,linuxapp.cpp)



Classes:

TheOS	(implemented on different os level, static not abstract)
TheApp	("")

TheConfigFile (superclass, some abstract member func)
TheIniFile
TheRegistry	(Win only) 
TheBinaryFile

TheFileSystem
TheWinFileSystem	(half way done)
TheMacFileSystem	(not done)
TheLinuxFileSystem	(not done)
TheCompressedFileSystem (not done)
TheZipFileSystem (not done)

TheFile (uses above)



Todo:
SOCKS class (look for 3rd party/CORBA/WX)

TheIO:: or 
TheGameController
TheKeyboard
TheMouse

