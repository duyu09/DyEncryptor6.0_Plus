#include<iostream>
#include<fstream>
#include<cstdlib>
#include<cstdio>
#include<io.h>
#include<sstream>
using namespace std;
int main(int argc,char *argv[])
{
    system("echo Copyright 2021 by DUYU.");
    system("echo This is DyEnc File Destroying Module");
    if (argc<2)
    {
    	system("echo Argument Error. 2 Argument Needed.");
    	system("echo Argument1: File Path.");
    	system("echo Argument2: File Size.");
    }
    ofstream out(argv[1],ios::binary);
    if (!out.is_open()) 
	{
        cout << "Error Open File.  " << endl;
        exit(1);
    }
    
	    int fgh;int jdm;
		char buf[2048];
		for(int fxl=0;fxl<=2049;fxl=fxl+1)
		{
			buf[fxl]=' ';
		}
		jdm=(int)(atoi(argv[2])/2048)+1;
		for(fgh=1;fgh<=jdm;fgh=fgh+1)
		{
			out.write(buf,2048);
		}
	out.close();
	return 0;
}


