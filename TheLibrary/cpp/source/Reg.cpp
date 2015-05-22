#ifndef __ZPROP_CPP__
#define __ZPROP_CPP__

#include "the.h"
#include <string.h>
#include <windows.h>
#include <stdlib.h>
#include <assert.h>

//==============================================================================
//==============================================================================
TheRegistry::TheRegistry(): _hRegKey(0),_rootKey(HKEY_CURRENT_CONFIG) {
	_fileName="WinRegistry";
}

//==============================================================================
//==============================================================================
TheRegistry::~TheRegistry() {
	close();
}
//==============================================================================
//	open ( )
//==============================================================================
bool TheRegistry::open (const string& filename) {
	close();			// Close any Open registry key first.  Important.
	if (TheString::icmp(filename,"HKEY_CLASSES_ROOT")==0) {
		_rootKey=HKEY_CLASSES_ROOT;
	}
	else if (TheString::icmp(filename,"HKEY_CURRENT_USER")==0) {
		_rootKey=HKEY_CURRENT_USER;
	}
	else if (TheString::icmp(filename,"HKEY_USERS")==0) {
		_rootKey=HKEY_USERS;
	}
	else if (TheString::icmp(filename,"HKEY_LOCAL_MACHINE")==0) {
		_rootKey=HKEY_LOCAL_MACHINE;
	}
	else if (TheString::icmp(filename,"HKEY_CURRENT_CONFIG")==0) {
		_rootKey=HKEY_CURRENT_CONFIG;
	}
	else if (TheString::icmp(filename,"HKEY_DYN_DATA")==0) {
		_rootKey=HKEY_DYN_DATA;
	}
	else {
		assert(1);	// ERROR: illegal registry root name
		_rootKey=HKEY_LOCAL_MACHINE;	// default value 
		return false;
	}
	return true;
}

//==============================================================================
//==============================================================================
bool TheRegistry::openConfig() {
	string appPath,
		drive,dir,fileName,ext;	// component
	_FileSystem->getAppPath(appPath);
	_FileSystem->splitPath(appPath,drive,dir,fileName,ext);
	string key="software\\"+fileName;
	open("HKEY_LOCAL_MACHINE");
	return openGroup(key,true);
}

//==============================================================================
//==============================================================================
bool TheRegistry::openSystem(const string& filename) {
	open("HKEY_LOCAL_MACHINE");
	return openGroup(filename);

}

//==============================================================================
//==============================================================================
bool TheRegistry::openUninstallKey(const string& productName) {
	close();
	string key="software\\microsoft\\windows\\currentVersion\\uninstall\\";
	key=key+  productName;
	open("HKEY_LOCAL_MACHINE");
	return openGroup(key);
}

//==============================================================================
//	close ( )
//==============================================================================
bool TheRegistry::close()
{
	closeGroup();
	_hRegKey = 0;
	_rootKey=0;
	_currentGroup="";
	return true;
}

//==============================================================================
//	flush()
//			It is recommended not to call this.  It is here just for a reference
//==============================================================================
bool TheRegistry::flush() {
	if (_hRegKey) {
		int result=RegFlushKey(_hRegKey);
		return ((bool)(result==ERROR_SUCCESS)); 
	}
	return false;
}

//==============================================================================
//	openGroup ( )
//==============================================================================
bool TheRegistry::openGroup(const string& group, bool create)
{
	// first, switch any forward slash to backslash as WinAPI won't see it right.
	string newGroup=group;
	TheString::replaceForwardSlashWithBackSlash(newGroup);
	// it is already opened.  
	if (_currentGroup==newGroup && _hRegKey) {
		return true;
	}
	closeGroup();		// close any old one first.

	// CREATE mode: If it exists, open it. If not, Create it.
	if (create) {
		DWORD result;
		RegCreateKeyEx(_rootKey,newGroup.c_str(),0,"",REG_OPTION_NON_VOLATILE,
			KEY_ALL_ACCESS,NULL,&_hRegKey, &result);
			_currentGroup=newGroup;
			return true;
	}
	// OPEN mode: open only if it exists.
	// open Group if it is new.
	if (RegOpenKeyEx( _rootKey, newGroup.c_str(), 0, KEY_ALL_ACCESS, &_hRegKey)
		!=ERROR_SUCCESS) {
		// key doesn't exist on the registry.
		_currentGroup="";
		return false;
	}
	_currentGroup=newGroup;
	return true;
}

//==============================================================================
//	closeKey ( )
//==============================================================================
void TheRegistry::closeGroup()
{
	if (_hRegKey) {
		RegCloseKey (_hRegKey);
		_currentGroup="";
		_hRegKey=0;
	}

}

//==============================================================================
//==============================================================================
void TheRegistry::readStr(string& str, const string& group, const string& key, 
						  const string& defaultString) {
	unsigned long valueType, dataSize;		//  temporary value
	char data[MAX_STRING];	// holds retrieved Value given group/key
	dataSize = MAX_STRING;

	assert(_rootKey);


	// open reg
	if (!openGroup(group)) {
		assert(1);	// error!  Can't open reg for some reason.
		str=defaultString;
		return;
	}

	// read reg value
	if ( RegQueryValueEx(_hRegKey, key.c_str(), 0, &valueType,
		(unsigned char*) data, &dataSize) !=ERROR_SUCCESS) {
		str=defaultString;	// KEY not found.
		return;
	}

	if ((valueType == REG_SZ) || (valueType == REG_EXPAND_SZ)) {
		str= data;
	}
	else if (valueType == REG_MULTI_SZ) {
		char* dataPtr=data;
		str="";
		while (*dataPtr!=0) {
			str=str+dataPtr+"\n";	// get next line.
			dataPtr=dataPtr+strlen(dataPtr);
		}
	}
	else if(valueType == REG_DWORD || valueType == REG_DWORD_LITTLE_ENDIAN) {
		// convert # to string.	
		long* numberPtr = (long*)data;
		char numberStr[32];
		sprintf (numberStr,"%d",*numberPtr);
		str=numberStr;	
	}
	else { 
		str=defaultString;		// non suported type
	}
}

//==============================================================================
//==============================================================================
bool TheRegistry::write(const string& group, const string& key,
						const string& value) {
	DWORD result=0;

	// if Group key doesn't exist in registry, create it first.
	openGroup(group, true);

	// write INTEGER
	if (TheString::isInteger(value)) {
		DWORD number=atoi(value.c_str());
		result= RegSetValueEx (_hRegKey, key.c_str(), 0, REG_DWORD, 
			(const unsigned char*)number,	4 );	// assuming dword is 4 byte wide (32bit).
	}
	// write STRING
	else {
		result= RegSetValueEx (_hRegKey, key.c_str(), 0, REG_SZ, 
			(const unsigned char*)value.c_str(),	value.length() );
	}		
	return (result==ERROR_SUCCESS);
}

//==============================================================================
//
//	By Default,
//	on NT, if it has subkey, it won't delete the group key.
//	on Win9X, it will delete the group key no matter what.
//	So it needs to TRAVERSE and delete the registry.  Be careful!
//==============================================================================
bool TheRegistry::deleteGroup(const string& group) {
	int result,i;

	// open group
	string newGroup=group;
	TheString::replaceForwardSlashWithBackSlash(newGroup);
	openGroup(newGroup);	
	if (!result) {		// ERROR. cannot delete what doesn't exist.
		return false;
	}

	// delete all subkeys.
	DWORD subKeySize=MAX_STRING, subKeyType,subKeyValueSize=MAX_STRING;
	char subKey[MAX_STRING];
	unsigned char subKeyValue[MAX_STRING];	// 
	for (i=0, result=ERROR_SUCCESS; result==ERROR_SUCCESS; i++) {
		result=RegEnumValue(_hRegKey, i, subKey, &subKeySize, NULL,
			&subKeyType,subKeyValue, &subKeyValueSize);
		if (result==ERROR_SUCCESS) {
			deleteKey(newGroup,subKey);
		}
	}

	// delete all subGroups
	char subGroup[MAX_STRING], subGroupClass[MAX_STRING];
	DWORD subGroupSize=MAX_STRING, subGroupClassSize=MAX_STRING;
	FILETIME lastWrittenTo;

	for (i=0, result=ERROR_SUCCESS; result==ERROR_SUCCESS; i++) {
		result=RegEnumKeyEx(_hRegKey, i, subGroup, &subGroupSize, NULL,
			subGroupClass, &subGroupClassSize, &lastWrittenTo);
		if (result==ERROR_SUCCESS) {
			deleteGroup(newGroup+"\\"+subGroup);
		}
	}
	// delete itself.
	openGroup(newGroup);	// make sure to reopen the group since _hRegKey was 					
							// altered by recursive func.
	result=RegDeleteKey(_hRegKey, newGroup.c_str());
	return ((bool)(result==ERROR_SUCCESS));
	closeGroup();
}

//==============================================================================
//==============================================================================
bool TheRegistry::deleteKey(const string&  group, const string& key) {
	openGroup(group);	
	int result=RegDeleteValue(_hRegKey, key.c_str());
	return ((bool)(result==ERROR_SUCCESS));
}


#endif