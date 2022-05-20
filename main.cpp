/*
author:GJD
date:2022-05-20
description:This program is used to split excel tables.
*/
#include <OpenXLSX.hpp>
#include <iostream>
#include <string.h>

using namespace std;
using namespace OpenXLSX;
char filename[100];
char noDot[100];
char tempfile[100];
int row,col;
int now=2;
int tag=1;
int main()
{
    int split;
    int row,col;
    int now=2;
    cout << "********************************************************************************\n";
    cout <<"Author:GJD"<<endl;
    cout<<"Date:2022-05-20"<<endl;
    cout<<"email:jiadongguo57@gmail.com"<<endl;
    cout << "********************************************************************************\n";
    XLDocument input;
    cout<<"input filename:";
    cin>>filename;
    input.open(filename);
    cout<<"input split lines:";
    cin>>split;
    sprintf(noDot,"%s",filename);
    int n=strlen(noDot)-1;
    while(n>=0&&noDot[n]!='.')
    {
        n--;
    }
    if(n>=0)
    {
        noDot[n]='\0';
    }
    XLWorksheet ips=input.workbook().worksheet("Sheet1");
    row=ips.rowCount();
    col=ips.columnCount();
    //XLDocument output;
    XLDocument *output;
    //XLWorksheet ops;
    XLWorksheet *ops;
    for(int i=2;i<=row;i++)
    {
        //whether to create a new file
        if(i==2)
        {
            sprintf(tempfile,"%s%d.xlsx",noDot,tag);
            tag++;
            output=new XLDocument;
            output->create(tempfile);
            ops=new XLWorksheet;
            *ops=output->workbook().worksheet("Sheet1");
            for(int k=1;k<=col;k++)
            {
                ops->cell(1,k).value()=ips.cell(1,k).value();
            }
            now=2;
        }
        else if(now==split+2)
        {
            cout<<tempfile<<" success!!"<<endl;
            output->save();
            output->close();
            delete output;
            delete ops;
            sprintf(tempfile,"%s%d.xlsx",noDot,tag);
            output=new XLDocument;
            output->create(tempfile);
            ops=new XLWorksheet;
            *ops=output->workbook().worksheet("Sheet1");
            for(int k=1;k<=col;k++)
            {
                ops->cell(1,k).value()=ips.cell(1,k).value();
            }
            now=2;
            tag++;
        }
        for(int j=1;j<=col;j++)
        {
            ops->cell(now,j).value()=ips.cell(i,j).value();
        }
        now++;
    }
    output->save();
    output->close();
    input.save();
    cout<<tempfile<<" success!!"<<endl;
    delete ops;
    delete output;
    cout<<"print any key to exit..."<<endl;
    getchar();
    getchar();
    return 0;
}