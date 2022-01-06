 //CopyRight 2021 by DUYU.  ���������Ȩ����
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
    
    copy(argv[1], argv[2], x, head2, atoi(argv[5]));   //argv[5]����ģʽ��0 = ���ܣ�1 = ���ܡ�
    

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
            //read��in���ж�ȡ2048�ֽڣ�����buf�����У�ͬʱ�ļ�ָ������ƶ�2048�ֽ�
            //������2048�ֽ������ļ���β������ʵ����ȡ�ֽڶ�ȡ��
            in.read(buf, 2048);
            //gcount()������ȡ��ȡ���ֽ�����write��buf�е�����д��out����
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
            //read��in���ж�ȡ2048�ֽڣ�����buf�����У�ͬʱ�ļ�ָ������ƶ�2048�ֽ�
            //������2048�ֽ������ļ���β������ʵ����ȡ�ֽڶ�ȡ��
            in.read(buf, 2048);
            //gcount()������ȡ��ȡ���ֽ�����write��buf�е�����д��out����
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
