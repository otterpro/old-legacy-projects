#ifndef _CONFIGFILE_H_
#define _CONFIGFILE_H_

//=======================================================
//	ConfigFile Base
//=======================================================
class TheConfigFile {
protected:
	string _fileName;	// full filename or registry location
public:
	TheConfigFile() {}
	~TheConfigFile() {}
//	TheConfigFile(const string& fileName="")  : _fileName(fileName) {
//		if (fileName!="") open(fileName);
//	}
//	~TheConfigFile() { close();}

	virtual bool open (const string& filename)=0;		
	// opens ini file from where app is located.
	virtual bool close()=0;

	virtual bool openSystem(const string& filename)=0;	// Looks for this ini file in
						// Win: "C:\windows\" (or windows dir)
						// Mac: "Preferences Folder"
						// Linux: "\etc\"
	
	virtual bool openConfig()=0;			//	opens app's ini file ie
						//	"myApp.exe" -> opens "myApp.INI"

	virtual void readStr(string& str, const string& group, const string& key, 
							const string& defaultString="")=0;
	char* readStr(const string& group, const string& key, char* buffer, int bufferSize,
					const string& defaultString="") {
					string str;
					readStr(str,group,key,defaultString);
					strcpy (buffer, str.c_str());
					return buffer;
				}

	int readInt(const string& group, const string& key, int defaultValue=0) {
					string str;
					readStr(str,group,key,"");
					if (str=="") return defaultValue;
					return atoi(str.c_str());
				};
	virtual bool write(const string& group, const string& key,const string& value)=0;
	bool write(const string& group, const string& key, int value) {
				char tempStr[32];
				return write(group,key,itoa(value,tempStr,10));
	}
	virtual bool deleteKey(const string&  group, const string& key)=0;
		// delete entire key and all its values
	virtual bool deleteGroup(const string& group)=0;
		// delete entire section, key.  
	virtual bool flush()=0;
};
#endif