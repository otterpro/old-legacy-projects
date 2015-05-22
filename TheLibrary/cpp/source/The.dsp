# Microsoft Developer Studio Project File - Name="The" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Static Library" 0x0104

CFG=The - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "The.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "The.mak" CFG="The - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "The - Win32 Release" (based on "Win32 (x86) Static Library")
!MESSAGE "The - Win32 Debug" (based on "Win32 (x86) Static Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
RSC=rc.exe

!IF  "$(CFG)" == "The - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "..\lib\release"
# PROP Intermediate_Dir "..\lib\release"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_MBCS" /D "_LIB" /YX /FD /c
# ADD CPP /nologo /W3 /GX /O2 /I "..\include\\" /D "WIN32" /D "NDEBUG" /D "_MBCS" /D "_LIB" /YX /FD /c
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LIB32=link.exe -lib
# ADD BASE LIB32 /nologo
# ADD LIB32 /nologo /out:"..\lib\The.lib"

!ELSEIF  "$(CFG)" == "The - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "The___Win32_Debug"
# PROP BASE Intermediate_Dir "The___Win32_Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "..\lib\debug"
# PROP Intermediate_Dir "..\lib\debug"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_MBCS" /D "_LIB" /YX /FD /GZ /c
# ADD CPP /nologo /W3 /Gm /GX /ZI /Od /I "..\include\\" /D "WIN32" /D "_DEBUG" /D "_MBCS" /D "_LIB" /YX /FD /GZ /c
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LIB32=link.exe -lib
# ADD BASE LIB32 /nologo
# ADD LIB32 /nologo /out:"..\lib\Thed.lib"

!ENDIF 

# Begin Target

# Name "The - Win32 Release"
# Name "The - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\CompressedFileSystem.cpp
# End Source File
# Begin Source File

SOURCE=.\FileSystem.cpp
# End Source File
# Begin Source File

SOURCE=.\IniFile.cpp
# End Source File
# Begin Source File

SOURCE=.\Misc.cpp
# End Source File
# Begin Source File

SOURCE=.\MMEDIA.CPP
# End Source File
# Begin Source File

SOURCE=.\Reg.cpp
# End Source File
# Begin Source File

SOURCE=.\STRING.CPP
# End Source File
# Begin Source File

SOURCE=.\The.cpp
# End Source File
# Begin Source File

SOURCE=.\winfilesystem.cpp
# End Source File
# Begin Source File

SOURCE=.\WINTOOL.CPP
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=..\include\CompressedFileSystem.h
# End Source File
# Begin Source File

SOURCE=..\include\configfile.h
# End Source File
# Begin Source File

SOURCE=..\include\FileSystem.h
# End Source File
# Begin Source File

SOURCE=..\include\IniFile.h
# End Source File
# Begin Source File

SOURCE=..\include\Misc.h
# End Source File
# Begin Source File

SOURCE=..\include\MMEDIA.H
# End Source File
# Begin Source File

SOURCE=..\include\Reg.h
# End Source File
# Begin Source File

SOURCE=..\include\StringVector.h
# End Source File
# Begin Source File

SOURCE=..\include\THE.H
# End Source File
# Begin Source File

SOURCE=..\include\winfilesystem.h
# End Source File
# End Group
# Begin Source File

SOURCE=.\readme.txt
# End Source File
# End Target
# End Project
