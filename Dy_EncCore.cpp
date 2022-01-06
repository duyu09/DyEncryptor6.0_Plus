 //CopyRight 2021 by DUYU.  杜宇保留所有权利。
 //All Rights Reserved.
 //2021-08-29.

#include<iostream>
#include<fstream>
#include<cstdlib>
#include<cstdio>
#include<io.h>

void copy(char* src, char* dst, int key, int headt, int mdw);
int main(int argc,char *argv[])
{
    using namespace std;
    system("echo Copyright 2021 by DUYU.");
    system("echo ***********************");
    
    if (argc<5)
    {
    	system("echo Arguement Error. 5 Arguements Needed.");
    	system("echo Arguement1: Source File.");
    	system("echo Arguement2: Destination File.");
    	system("echo Arguement3: KEY.");
		system("echo Arguement4: Length of File Head.");
		system("echo Arguement5: Running Mode,0=Encrypt,1=Unencrypt.");
    	system("echo For Example:");
    	system("echo Dy_EncCore.exe D:\\input.txt D:\\output.txt 9 100 0");		
    	exit(1);
	}
    
    int head2 = atoi(argv[4]);
    int x = atoi(argv[3]);
    
    copy(argv[1], argv[2], x, head2, atoi(argv[5]));   //argv[5]表明模式，0 = 加密，1 = 解密。
    

    return 0;
}

void copy(char* src, char* dst, int key, int headt, int mdw)
{   
	if (mdw == 0)
	{
		using namespace std;
        ifstream in(src,ios::binary);
        ofstream out(dst,ios::binary);
        if (!in.is_open()) 
		{
            cout << "Error Open File.  " << src << endl;
            exit(1);
        }
        if (!out.is_open()) 
		{
            cout << "Error Open File.  " << dst << endl;
            exit(1);
        }
        if (src == dst)
		{
            cout << "Source File Can't Be Same With Destination File." << endl;
            exit(1);
        }
        char buf[2048];
        long long totalBytes = headt;
        out.clear();
        out.seekp(headt,ios::beg);
        
        while(in)
        {
            //read从in流中读取2048字节，放入buf数组中，同时文件指针向后移动2048字节
            //若不足2048字节遇到文件结尾，则以实际提取字节读取。
            in.read(buf, 2048);
            //gcount()用来提取读取的字节数，write将buf中的内容写入out流。
            char buft[2048];
            for(int a=0;a<=2048;a=a+1)
            {
        	    buft[a]=(char)((int)buf[a]^(int)key);
		    }
            out.write(buft, in.gcount());
            totalBytes += in.gcount();
        }
        in.close();
        out.close();
            system("echo finish >> \"%tmp%\\finish.dyenc\"");
            system("echo Encrypted Successfully.");
            cout << "Total Bytes: " << totalBytes << " B" <<endl;
        exit(0);
    }
    
    
    else
	{
		using namespace std;
        ifstream in(src,ios::binary);
        ofstream out(dst,ios::binary);
        if (!in.is_open()) 
		{
            cout << "Error Open File.  " << src << endl;
            exit(1);
        }
        if (!out.is_open()) 
		{
            cout << "Error Open File.  " << dst << endl;
            exit(1);
        }
        if (src == dst)
		{
            cout << "Source File Can't Be Same With Destination File." << endl;
            exit(1);
        }
        char buf[2048];
        long long totalBytes = headt;
        in.clear();
        in.seekg(headt,ios::beg);
        while(in)
        {
            //read从in流中读取2048字节，放入buf数组中，同时文件指针向后移动2048字节
            //若不足2048字节遇到文件结尾，则以实际提取字节读取。
            in.read(buf, 2048);
            //gcount()用来提取读取的字节数，write将buf中的内容写入out流。
            char buft[2048];
            for(int a=0;a<=2048;a=a+1)
            {
        	    buft[a]=(char)((int)buf[a]^(int)key);
		    }
            out.write(buft, in.gcount());
            totalBytes += in.gcount();
        }
        in.close();
        out.close();
            system("echo finish >> \"%tmp%\\finish.dyenc\"");
            system("echo Encrypted Successfully.");
            cout << "Total Bytes: " << totalBytes << " B" <<endl;
        exit(0);
    }
}
