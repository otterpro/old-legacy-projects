//==============================================================================
//	MISC.H
//	
//	Generic and common low-level misc classes and functions.  They don't depend
//		on anything except for standard library
//==============================================================================
#ifndef __MISC_H_
#define __MISC_H_

#define MAX_STRING 255
#define MAX_DIGIT 32


#ifdef WIN32
#define _PATH_SEPARATOR_CHAR '\\'
#define _PATH_SEPARATOR_STRING "\\"
#else
#define _PATH_SEPARATOR_CHAR ':'
#define _PATH_SEPARATOR_STRING ":"
#endif


#include <time.h>
#include <string>
using namespace std;
#include <string.h>

// Strip leading and trailing spaces
void stripWhiteSpace(string& str);


//==============================================================================
// Rectangle class
//==============================================================================
class TheRect {
public:
	int _x1, _x2, _y1, _y2;	// 
	TheRect(int x1=0, int x2=0, int y1=0, int y2=0) :_x1(x1),_y1(y1),_x2(x2),_y2(y2) {}
	bool equal(const TheRect& rect) {
		if (rect._x1==_x1 && rect._x2==_x2 && rect._y1==_y1 && rect._y2==_y2) {
			return true;
		}
		return false;
	}
	void copy(const TheRect& rect) {
		_x1=rect._x1;
		_x2=rect._x2;
		_y1=rect._y1;
		_y2=rect._y2;
	}
};

//==============================================================================
// Coordinate (xy) class
//==============================================================================
class TheXY {
public:
	int _x, _y;
	TheXY(int x=0, int y=0) : _x(x), _y(y) {}
	bool equal(const TheXY& xy) {
		if (xy._x==_x && xy._y==_y) { return true;}
		return false;
	}
	void copy(const TheXY& xy) {
		_x=xy._x;
		_y=xy._y;
	}
	void set(const string& xyStr) {		// converts string "x,y" to _x, _y.
		char str[MAX_STRING], *x, *y;
		strcpy(str,xyStr.c_str());
		x=strtok(str,",");
		y=strtok(NULL,",");
		_x=atoi(x);
		_y=atoi(y);
	}
};

//==============================================================================
//	String	
//				Utility for std "string" 
//==============================================================================
class TheString {
public:
									
	static int icmp(const string& strA,const string& strB) {	
								return stricmp(strA.c_str(),strB.c_str());
								}			// stricmp()
	static int nicmp(const string& strA,const string& strB, int length) {
								return strnicmp(strA.c_str(),strB.c_str(),length);
								}			// strnicmp	
	static void addIniBracket(string& str);		// surround a string with 
												// "[" and "]"  if needed.
	static void removeIniBracket(string& str);	// remove 1st "[" and "]" 
	static void removeComment(string& str);	// remove all whitespace and comment 
											// ("; and #");
	static void removeLeadingWhiteSpace(string& str); // remove "___blah blah"
	static void removeTrailingWhiteSpace(string& str);	// remove "blah blah___"

	static void toLower(string& str);		// make into lowercase
	static void toUpper(string& str);		// make into uppercase

	static bool isInteger(const string& str);	// returns true if given string
											// is Integer.  Can't read sci notation
	static void replaceForwardSlashWithBackSlash(string& str);
							// make all "/" into "\\". Useful for PC/Win
	static bool isSectionName(const string& section);
							// returns true if it is a section 
							// (ie "[blah]" => true,  "blah"=>false)
									
	static string& itoa(int value);
	static string& ftoa(float value);
	//friend string& operator+(const string& str, int number); 
	static string& concat(const string& str, int number);
							// Warning:  Does not handle NEGATIVE Numbers (ie -1, -3, -5.03)
							// ex:  ("sprite",1) => returns "sprite1" 
							//		("sprite",3) => returns "sprite3"
							//		("sprite",-1)=> returns "sprite".
	static string& concat(const string& str1, const string& str2);
};

//==============================================================================
// process and parses command line argument
//==============================================================================
class TheCmdLineArg {
protected:
	string _cmdLineArgStr;	// entire cmdLineArg
	int _cmdLineArgIndex;	// n-th cmd line arg param
	//StringVector _cmdLineArg;	
public:
	TheCmdLineArg(const char* cmdLineArg=0);
	virtual ~TheCmdLineArg();
	int open(const char* cmdLineArg);
	int close();
	string getNext();	// get next cmdline string. Returns "" if no more cmdline string.
	string getAll();	// show entire cmdlinearg as a string
	string get(int index);	// get n-th cmdlinearg. 
};

//==============================================================================
//	Slow Random Func
//==============================================================================
class TheRandom {
public:
	TheRandom() {
		init();
	}
	static void init() {
		srand(time(0));
	}
	static int get(int min, int max) {
		int range=max-min;	
		int rawRandomValue=rand()%(range+1);
		return (rawRandomValue+min);
	}
	static int get5050() {
		return (rand()%2);
	}
};

//==============================================================================
//	Index
//==============================================================================
class TheIndex {
public:
	TheIndex(int min=0, int max=255);
	int get();
protected:
	int _currentIndexNumber,
		_min, _max;
};

//==============================================================================
//==============================================================================
int parser(unsigned inflag,char* token,int tokmax,char* line,char* white,
		   char* brkchar,char* quote, char eschar,char* brkused,int* next,
		   char* quoted);

/* Parser() example:
	char whitesp[]={" \t"};	/// blank and tab 
	char breakch[]={""};	// comma and carriage return 
	char quotech[]={"'\""};	// single and double quote 
	char escape[]={"^"};		// "uparrow" is escape 

	  char line[81],brkused,quoted,token[81];
	  int i,next;
	strcpy (line, cmdline);

	    //printf("Line: %s",cmdline);		// already has <CR> 
	    i=0;

	    next=0;				// make sure you do this 

	    while(parser(2,token,80,line,whitesp,breakch,quotech,escape,
			 &brkused,&next,&quoted)==0)
	    {
	      //printf(" Token %d = (%s)\n",++i,token);
			
			if (stricmp (token, "-test")==0) 
			{
				bTest = TRUE;
			}
			if  (stricmp (token,"-noreboot") ==0) 
			{
				bReboot= FALSE;
			}

	      if(brkused=='\r')	// <CR> is a break so it won't be included  
		break;		// in the token.  treat as end-of-line here 
	    }
*/
#endif
