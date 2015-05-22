#ifndef _THE_REG_H_
#define _THE_REG_H_

#include "the.h"
#include <windows.h>

#include <string>
using namespace std;

//==============================================================================
//	TheRegistry
//	
//	Reggie:  reg manager
//
//	BUG:
//		On WinNT/2000 Prof, the "deleteGroup()" doesn't delete all the contents
//			of the group.  Tried recursive call but still not deleting itself.
//
//==============================================================================
class TheRegistry : public TheConfigFile {
	
public:
	TheRegistry();
	~TheRegistry();
	virtual bool open (const string& filename);		// opens registry value.
						// filename: either HKEY_LOCAL_MACHINE, HKEY_CLASSES_ROOT,
						//			HKEY_CURRENT_USER, HKEY_USERS, 
						//			HKEY_CURRENT_CONFIG, HKEY_DYN_DATA

	virtual bool openConfig();	//	opens app's default registry point
								//	ie: (running ThisApp.exe)
								//	HKEY_LOCAL_MACHINE\software\ThisApp\.
								// if Key doesn't exist, it WILL create it.

	virtual bool openSystem(const string& filename);	
								// ex: openSystem("BLAH") ->
								// opens HKEY_LOCAL_MACHINE\software\BLAH
								// if key doesn't exit, it WON'T Create it.
	virtual void readStr(string& str, const string& group, const string& key, 
							const string& defaultString="");
	virtual bool write(const string& group, const string& key,
						const string& value);
	virtual bool deleteKey(const string&  group, const string& key);
							// delete entire key and all its values
	virtual bool deleteGroup(const string& group);
							// delete entire section, key.  
	virtual bool flush();	// Try not to use this in registry.
	virtual bool close();

	// MISC 
	bool openUninstallKey(const string& productName);
							// open uninstall key immediately.
							// if key doesn't exist, it WON'T create it.
	// debug function
	void dump();		// show the current content of INI file to screen.

protected:
	HKEY _hRegKey;			// handle to registry key

	HKEY _rootKey;			// HKEY_LOCAL_MACHINE
	string _currentGroup;	// currently selected group ie [section]
							// ex: "software/microsoft/blah"

	bool openGroup(const string& group, bool create=false);
							// open the group (ie registry key). It is cached
							// so that it doesn't reopen the same group if 
							// it is opened from previous operation.
							// "create" -> if set to TRUE, it will create the
							//			[group] if it doesn't already exist.
							//			Always returns TRUE
							// "create"-> if set to FALSE, 
							// Returns FALSE if key doesn't exist. Typically,
							// write() will get false since key may not exist.

	void closeGroup();		// close any opened group. Used only by openGroup().

};

/*
class Registry {
protected:
	HKEY _hRegKey;			// handle to registry key
	string _subKey;	// "Software/Microsoft/..."
public:
	//Registry ();  
	~Registry ( ) { close (); }
	int open(const char* key, const char* rootKey="HKEY_LOCAL_MACHINE");

	int openUninstallKey(const char* productKey);
	// goes to HKEY_Local_machine\software\microsoft\windows\currentVersion\
	//			\uninstall\productKey

	void close ();		
	
	string getStr(const char* valueName);	
	
	// NOT tested YET.
	int getInt(const char* valueName, int defaultValue=0);	
	float getFloat(const char* valueName, float defaultValue=0.0);	

	
	int set (const char* valueName, const char* valueData);
	// not implemented yet
	int set (const char* valueName, int valueData);
	int set (const char* valueName, float valueData);

	int deleteValue(const char* valueName);
	int deleteKey();	// closes current key and delete it. Make sure you
		// open() again if you want to keep using the same reg class 
	int deleteKey(const char* subKey);	// delete a subkey of current key.
};
*/
#endif		// _zprop.h//