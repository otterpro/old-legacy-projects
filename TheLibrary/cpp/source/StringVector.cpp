#ifndef __ZMISC_CPP__
#define __ZMISC_CPP__

#include <stdlib.h>
#include <string.h>
#include "StringVector.h"

//===================================
//	zConsole 
//
//	a base class that keeps track of a fixed-width
//	character based console.  (Ex: 80x24 terminal).
//	This is based on Quake Console that keeps track
//	of basic user commands.
//===================================
//-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
//
//-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
void zConsoleBuffer::init (int x, int y, int line)
{	   
//	static char _mainText[128][128];
	maxWidth = x; maxHeight = y; maxLine = line;
	if ( maxLine < maxHeight ) maxLine = maxHeight;	// error case

	currentX = 0; 
	currentY = 0;
	currentLine = 0;

	//outputText = ( char* ) malloc ( maxWidth * maxHeight + maxHeight);
//	mainText = _mainText;
	//mainText = ( char* ) malloc ( maxWidth * (maxLine + 1));
	
	// add maxHeight / maxLine at the end just to add a fudge factor

	
//	for (int i = 0; i < maxLine; i ++ )
//	{}

	cls();		// clear any content.

	for (int i = 0; i < ZCONSOLE_MAXLINE; i++)	// clear entire main buffer
	{
		mainText[i][0] = 0;
		//strcpy (mainText[i][0], "");
	}
}

////=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
///
//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
void zConsoleBuffer::scrollUp (int byline)		// scroll up by x lines.
{
	println ("scrollup() not supported yet.");
}

////=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
///
//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
void zConsoleBuffer::scrollDown (int byline)		// scroll down by x lines
{
	println ("scrolldown() not supported yet.");
}

////=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
///
//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
void zConsoleBuffer::setMaxLine (int line)			// sets maximum line (history)
{
	if (line >=ZCONSOLE_MAXLINE) 
	{
		line = ZCONSOLE_MAXLINE-1;
		println ("exceeded 128");
	}
	println ("setmaxline=x not supported yet.");
}

////=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
///
//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
void zConsoleBuffer::print (char* szText)
{
//	int i;
	//if (currentX >= ZCONSOLE_MAXWIDTH)
	// TODO:	chop off the string szText so that it will print partially
	//			instead of ignoring the entire string.
	if ((int)strlen (szText) + (int)strlen (mainText[currentLine]) >= maxWidth)
	{
		return;			//ERROR:  string is too long to print.
	}
	strcat (mainText[currentLine], szText);		// main buffer
//	strcat (outputText, szText);			// write to output temp buffer
}

////=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
///	println
//	
//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
void zConsoleBuffer::println (char* szText)
{
	print(szText);
//	print("\n");
	currentLine++;
	if (currentLine >= maxLine)			// scroll down one line.
	{
		for (int i = 0; i < maxLine-1; i++)
		{
			strcpy (mainText[i], mainText[i+1]);
		}
		mainText[maxLine-1][0] = 0;	// last line = empty
		currentLine --;
	}
	 
	//outputLine++;
//	if (outputLine >= maxHeight)
//	{
//		char* szTemp;
//		szTemp = strchr(outputText, '\n');	// find the NewLine marker
//		strcpy (outputText, szTemp+1);
//	}

	//advanceLine();
}

////=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
///
//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
void zConsoleBuffer::cls ()
{
//	outputLine=0;
	currentLine=0;
	firstLineNumber = 0;
//	lastLineNumber = 20;	// debugging purpose only.  REMOVE

	lastLineNumber = maxHeight;
//	outputText[0] = 0;	// clear output buffer
	for (int i = 0; i < maxLine; i++)	// clear main buffer
	{
		mainText[i][0] = 0;
		//strcpy (mainText[i][0], "");
	}

}

// AppendFilePath("Milo","def.txt"); will return "Milo\\def.txt" for PC and "Milo:def.txt" for MAC
string AppendFileToPath(const char* dir, const char* file) {
	string newPath;	// path to return.
	char newPathStr[1024];	// enough to hold the path string.
	strcpy (newPathStr, dir);
#ifdef _MAC
	strcat (newPathStr, ":");
#else
	strcat (newPathStr, "\\");
#endif
	strcat (newPathStr, file);
	newPath = newPathStr;
	return newPath;
}

// WARNING: not a thorough func.  Use only "/" forward slash!
string GetFileNameFromURL(const char* url) {
	string filename;
	char filenameStr[1024];	// enough to hold the path string.
	// if it ends in "/", it is an error (ie no file is found)
	if (url[strlen(url)-1] == '/') {
		return "";
	}

	int filenamePosition=0;	// position where filename in the url begins.
	int i = strlen(url)-1;
	do {
		if (url[i] == '/') {
			filenamePosition = i + 1;
		}
		i--;
	}while (i>=0 && !filenamePosition);
	//char* fileNamePositionPtr = url[filenamePosition+1];
	
	strcpy(filenameStr,&url[filenamePosition]);
	filename = filenameStr;
	return filename;
}
#endif
//__ZMISC_CPP__