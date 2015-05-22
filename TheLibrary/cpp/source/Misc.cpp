#include "misc.h"
#include <stdlib.h>
#include <ctype.h>

//________________________________________________________________________________
//	AppendFilePath("Milo","def.txt"); 
//	will return "Milo\\def.txt" for PC and "Milo:def.txt" for MAC
//________________________________________________________________________________
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


//________________________________________________________________________________
//________________________________________________________________________________
void TheString::removeLeadingWhiteSpace(string& str) {
	string::iterator ptr=str.begin();
	if (!isspace(str[0])) return;	// no leading ws
	while ( ptr!=str.end() ) {
		if (isspace(*ptr)) {
			str.erase(ptr);
		}
		else {
			return;	// reached end of leading ws.
		}
		ptr++;
	}
}

//________________________________________________________________________________
//________________________________________________________________________________
void TheString::removeTrailingWhiteSpace(string& str) {
	string::iterator ptr=str.end();
	if (!isspace(str[0])) return;	// no leading ws
	while ( ptr!=str.begin() ) {
		if (isspace(*ptr)) {
			str.erase(ptr);
		}
		else {
			return;	// reached end of leading ws.
		}
		ptr--;
	}
}
//________________________________________________________________________________
//________________________________________________________________________________
void TheString::removeIniBracket(string& str) {
	if (isSectionName(str)) {
		str=str.substr(1,str.length()-2);	// remove "[" and "]"
	}
}

//________________________________________________________________________________
//________________________________________________________________________________
void TheString::addIniBracket(string& str) {
	if (isSectionName(str)) {
		return;		// already has "[" and "]"
	}
	str="["+str+"]";
}

//________________________________________________________________________________
//________________________________________________________________________________
void TheString::toLower(string& str) {
	int i=0;
	for (i=0; i<str.length(); i++) {
		str[i]=tolower(str[i]);
	}
}

//________________________________________________________________________________
//________________________________________________________________________________
void TheString::toUpper(string& str) {
	int i=0;
	for (i=0; i<str.length(); i++) {
		str[i]=toupper(str[i]);
	}
}

//________________________________________________________________________________
//	isInteger()
//________________________________________________________________________________
bool TheString::isInteger(const string& str) {
	for (int i=0; i < str.length(); i++) {
		if (!(isdigit(str[i]))) {
			return false;
		}
	}
	return true;
}

//________________________________________________________________________________
//________________________________________________________________________________
void TheString::replaceForwardSlashWithBackSlash(string& str) {
	int slashPosition=str.find('/');
	while (slashPosition!=string::npos) {
		str.replace(slashPosition,1,"\\");
		slashPosition=str.find('/');
	}
}

//________________________________________________________________________________
//________________________________________________________________________________
bool TheString::isSectionName(const string& section) {
	return (section[0]=='['); 
}

//________________________________________________________________________________
//________________________________________________________________________________
string& TheString::itoa(int value) {
	static string returnStr="";
	char numStr[MAX_DIGIT];
	::itoa(value,numStr,10);
	returnStr=numStr;
	return returnStr;
}
//________________________________________________________________________________
//________________________________________________________________________________
string& TheString::ftoa(float value) {
	static string returnStr="";
	char numStr[MAX_DIGIT];
	sprintf(numStr,"%f",value);
	returnStr=numStr;
	return returnStr;
}

//________________________________________________________________________________
//________________________________________________________________________________
string& TheString::concat(const string& str, int number) {
	// ex:  ("sprite",1) => returns "sprite1" 
	//		("sprite",3) => returns "sprite3"
	//		("sprite",-1)=> returns "sprite".
	static string newString="";
	newString=str;
	if (number>=0) {
		newString+=TheString::itoa(number);
	}
	return newString;
}


//________________________________________________________________________________
//________________________________________________________________________________
string& TheString::concat(const string& str1, const string& str2) {
	static string newString="";
	newString=str1+str2;
	return newString;
}

//________________________________________________________________________________
// WARNING: not a thorough func.  Use only "/" forward slash!
//________________________________________________________________________________
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

//________________________________________________________________________________
//________________________________________________________________________________

TheCmdLineArg::TheCmdLineArg(const char* cmdLineArg) {
	if (cmdLineArg) {
		open(cmdLineArg);
	}
}
TheCmdLineArg::~TheCmdLineArg() { 
	close();
}
int TheCmdLineArg::open(const char* cmdLineArg) {
	_cmdLineArgStr=cmdLineArg;
	_cmdLineArgIndex=0;	// n-th cmdline param.
	// clear cmdLine
	close();
	return 1;
}
int TheCmdLineArg::close() {
	//	_cmdLineArg.clear();
	return 1;
}
string TheCmdLineArg::getNext() {
	// parse it and return next str
	char whitesp[]={" \t"};	/* blank and tab */
	char breakch[]={""};	/* comma and carriage return */
	char quotech[]={"'\""};	/* single and double quote */
	char escape='^';		/* "uparrow" is escape */
	char line[MAX_STRING],brkused,quoted,token[MAX_STRING];
	string returnStr="";	// string to return
	strcpy (line, _cmdLineArgStr.c_str());

	if (parser(2,token,MAX_STRING,line,whitesp,breakch,quotech,escape,
		&brkused,&_cmdLineArgIndex,&quoted)==0)	{
		//printf(" Token %d = (%s)\n",++i,token);
		returnStr = token;
		//if(brkused=='\r')	/* <CR> is a break so it won't be included  */
		//	break;		/* in the token.  treat as end-of-line here */
	}
	return returnStr;
}

string TheCmdLineArg::getAll() {
	return _cmdLineArgStr;
}

//________________________________________________________________________________
//
//________________________________________________________________________________

TheIndex::TheIndex(int min, int max) : _min(min), _max(max), 
						_currentIndexNumber(0) {
}
int TheIndex::get() {
	return _currentIndexNumber++;
}

//________________________________________________________________________________
//
//________________________________________________________________________________
/* states */
//extern "C"{
#define IN_WHITE 0
#define IN_TOKEN 1
#define IN_QUOTE 2
#define IN_OZONE 3

int _p_state;      /* current state      */
unsigned _p_flag;  /* option flag        */
char _p_curquote;  /* current quote char */
int _p_tokpos;     /* current token pos  */

/* routine to find character in string ... used only by "parser" */

sindex(char ch,char* szString)
{
  char *cp;
  for(cp=szString;*cp;++cp)
    if(ch==*cp)
      return (int)(cp-szString);  /* return postion of character */
  return -1;                    /* eol ... no match found */
}
    
/* routine to store a character in a szString ... used only by "parser" */

void chstore(char* szString,int max,char ch) {
  char c;
  if(_p_tokpos>=0&&_p_tokpos<max-1)
  {
    if(_p_state==IN_QUOTE)
      c=ch;
    else
      switch(_p_flag&3)
      {
        case 1:             /* convert to upper */
          c=toupper(ch);
          break;
  
        case 2:             /* convert to lower */
          c=tolower(ch);
          break;
      
        default:            /* use as is */
          c=ch;
          break;
      }
    szString[_p_tokpos++]=c;
  }
  return;
}
  
/* here it is! */
int parser(unsigned inflag,char* token,int tokmax,char* line,char* white,char* brkchar,char* quote,
		   char eschar,char* brkused,int* next,char* quoted){
  int qp;
  char c,nc;
          
  *brkused=0;           /* initialize to null */	  
  *quoted=0;		/* assume not quoted  */

  if(!line[*next])      /* if we're at end of line, indicate such */
    return 1;

  _p_state=IN_WHITE;       /* initialize state */
  _p_curquote=0;           /* initialize previous quote char */
  _p_flag=inflag;          /* set option flag */

  for(_p_tokpos=0;c=line[*next];++(*next))      /* main loop */
  {
    if((qp=sindex(c,brkchar))>=0)  /* break */
    {
      switch(_p_state)
      {
        case IN_WHITE:          /* these are the same here ...	*/
        case IN_TOKEN:          /* ... just get out		*/
	case IN_OZONE:		/* ditto			*/
          ++(*next);
          *brkused=brkchar[qp];
          goto byebye;
        
        case IN_QUOTE:           /* just keep going */
          chstore(token,tokmax,c);
          break;
      }
    }
    else if((qp=sindex(c,quote))>=0)  /* quote */
    {
      switch(_p_state)
      {
        case IN_WHITE:   /* these are identical, */
          _p_state=IN_QUOTE;        /* change states   */
          _p_curquote=quote[qp];         /* save quote char */
          *quoted=1;	/* set to true as long as something is in quotes */
          break;
  
        case IN_QUOTE:
          if(quote[qp]==_p_curquote)	/* same as the beginning quote? */
	  {
            _p_state=IN_OZONE;
	    _p_curquote=0;
	  }
          else
            chstore(token,tokmax,c);	/* treat as regular char */
          break;

	case IN_TOKEN:
	case IN_OZONE:
	  *brkused=c;			/* uses quote as break char */
	  goto byebye;
      }
    }
    else if((qp=sindex(c,white))>=0)       /* white */
    {
      switch(_p_state)
      {
        case IN_WHITE:
	case IN_OZONE:
          break;		/* keep going */
          
        case IN_TOKEN:
          _p_state=IN_OZONE;
          break;
          
        case IN_QUOTE:
          chstore(token,tokmax,c);     /* it's valid here */
          break;
      }
    }
    else if(c==eschar)			/* escape */
    {
      nc=line[(*next)+1];
      if(nc==0)			/* end of line */
      {
	*brkused=0;
	chstore(token,tokmax,c);
	++(*next);
	goto byebye;
      }
      switch(_p_state)
      {
	case IN_WHITE:
	  --(*next);
	  _p_state=IN_TOKEN;
	  break;

	case IN_TOKEN:
	case IN_QUOTE:
	  ++(*next);
	  chstore(token,tokmax,nc);
	  break;

	case IN_OZONE:
	  goto byebye;
      }
    }
    else        /* anything else is just a real character */
    {
      switch(_p_state)
      {
        case IN_WHITE:
          _p_state=IN_TOKEN;        /* switch states */
          
        case IN_TOKEN:           /* these 2 are     */
        case IN_QUOTE:           /*  identical here */
          chstore(token,tokmax,c);
          break;

	case IN_OZONE:
	  goto byebye;
      }
    }
  }             /* end of main loop */

byebye:
  token[_p_tokpos]=0;   /* make sure token ends with EOS */
  
  return 0;
  
}

//}	// extern "C" 