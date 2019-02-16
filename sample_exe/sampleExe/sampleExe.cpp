
#pragma comment(lib, "xl_util_cpp.lib")

#include <string>

#include <iostream>

typedef std::string string;

// ここのプロトタイプ消す方
// ハンドル式にするt
__declspec(dllimport) int openWorkBook();
__declspec(dllimport) int addWorkSheet(string worksheet);
__declspec(dllimport) void setValue(string cell, string value,string worksheet);
__declspec(dllimport) std::string getValue(std::string cell, std::string worksheet);
__declspec(dllimport) int saveAs(std::string path);

int main()
{
	//std::string worksheetname;
	//std::string val;
	std::string retval;
	//std::string path;

	std::cout << "aiueo" << std::endl;
	std::cout << "as" << std::endl;
	std::cout << "a" << std::endl;
	std::cout << "WorkBookuname"<< std::endl;
	std::cout << "WorkSheetName" << std::endl;
	std::cout << "" << std::endl;



	printf( "aiueo\n" );

	//std::string path;

	std::cout << "nanngyoume" << std::endl;
	//std::cin >> path;


	//std::cout << path << std::endl;

	// 新規ワークブックの作成
	openWorkBook();

	// 
	addWorkSheet( "newWorkSheet" );

	// 
	setValue("newWorkSheet","A1", "aaa");
	
	// 
	retval = getValue("newWorkSheet", "A1");

	std::cout << retval << std::endl;


	// 
	saveAs("C:/Works/otameshi.xlsx");

	while(1){}

	getchar();

	return 0;
}
