#ifndef __STRING_VECTOR_H__
#define __STRING_VECTOR_H__

// StringVector is a header-only class since it is inlined for speed.

//#include <string>

#include <assert.h>
#include <stdio.h>

//#ifdef _MSC_VER					// MS VC++ Only
#include <vector>
#include <string>
using namespace std;
//#endif



//================================================================================
//
//	zStringVector
//
//	string array/vector.
//================================================================================
class StringVector
{
private:
	//vector<char*>strDB;				// array/vector of strings
	vector<string>strDB;				// array/vector of strings
public:
	int push_back (const string& str)
								{ 
										strDB.push_back (str);
										return size()-1;
										// Has to return size()-1 since size contains 1 more 
										// ex:  0 item, push -> size=1.  However, index of 1st item
										// is 0.  
									}
	int insert (int position, const string& str) {
										strDB.insert(strDB.begin()+position, str);
										return size()-1;
									}
	void erase(int position) {
								strDB.erase(strDB.begin()+position);
							}
	string& get(int i)			{
									assert (getSize()>=i);	// out of range error
									return strDB[i]; }
	void set(int i, const string& str) {
									assert(getSize()>=i);
									strDB[i]=str;
										}
	int size()					
									{ return getSize();}

	int getSize() 					
									{ return strDB.size();}
	StringVector ( ) {  }

	~StringVector( ) 
									{
										//char* oldStr;
										//for (int i = 0; i < strDB.size(); i ++)
										//{
										//	oldStr = strDB[i];
										//	free (oldStr);
										//}
									}
	int find(const char* _str)		// Search for a string in the database and return its index
	{							// or return -1 for error.  [should use -1?]
		for (int i = 0; i < size(); i ++)
			if (stricmp (_str, strDB[i].c_str())==0) return i;
		return -1;				// ERROR:  Not found.
	}
	// push_back only if it is unique in the string vector.  If it is not unique, return
	//	index # of the found item.
	int push_back_unique (const char* _str)
	{
		int i = find(_str);
		if (i==-1) 
		{		// Not found.  It is unique.
			return push_back (_str);
		}
		else
		{
			return i;
		}
	}
	void clear() { strDB.clear();}
};
#endif
//__ZMISC_H__
