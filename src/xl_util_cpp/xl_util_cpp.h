#pragma once

#if defined (_DEBUG)
#using "../xl_util_cs/bin/Debug/xl_util_cs.dll"
#else
#using "../xl_util_cs/bin/Release/xl_util_cs.dll"
#endif

#include <string>

// C#側に渡すための構造体
struct Data
{
	char path[64];
	char cell[64];
	char val_c[64];
	// int val_i;
};

// マネージドクラス
public ref class  XLUtil
{
private:
	xl_util_cs::WrapClosedXML work;

public:
	int openWorkBook()
	{
		return work.OpenWorkBook();
	}

	int addWorkSheet( std::string name )
	{
		Data data;

		// C#側に渡す構造体の作成
		strncpy_s( data.path, name.c_str(), name.length() );

		return work.AddWorkSheet( reinterpret_cast<int>(&data) );
	}

	int setValue( std::string ws, std::string cell, std::string val )
	{
		Data data;

		// C#側に渡す構造体の作成
		strncpy_s( data.path, ws.c_str(), ws.length() );
		strncpy_s( data.cell, cell.c_str(), cell.length() );
		strncpy_s( data.val_c, val.c_str(), val.length() );

		return work.SetValue( reinterpret_cast<int>(&data) );
	}

	std::string getValue( std::string ws, std::string cell )
	{
		Data data;

		// C#側に渡す構造体の作成
		strncpy_s( data.path, ws.c_str(), ws.length() );
		strncpy_s( data.cell, cell.c_str(), cell.length() );

		// 返ってきたデータポインタ
		int retdata = work.GetValue( reinterpret_cast<int>(&data) );
		if (retdata == NULL)
			return NULL;

		// 文字列を取得
		std::string ret( reinterpret_cast<Data*>(retdata)->val_c );

		return ret;
	}

	int saveAs( std::string path )
	{
		Data data;

		// C#側に渡す構造体の作成
		strncpy_s( data.path, path.c_str(), path.length() );

		return work.SaveAs( reinterpret_cast<int>(&data) );
	}
};

// 新規ワークブックの作成
__declspec(dllexport) int openWorkBook()
{
	XLUtil util;

	return util.openWorkBook();
}

// 新規ワークシートの追加
// name ワークシート名
__declspec(dllexport) int addWorkSheet(std::string name)
{
	XLUtil util;

	return util.addWorkSheet(name);
}

// 指定のセルに文字列を挿入
// ws	ワークシート
// cell	セル
// val	文字列
__declspec(dllexport) void setValue(std::string ws, std::string cell, std::string val)
{
	XLUtil util;

	util.setValue(ws, cell, val);
}

// 指定のセルの文字列を取得
// ws	ワークシート
// cell	セル
__declspec(dllexport) std::string getValue(std::string ws, std::string cell)
{
	XLUtil util;

	return util.getValue(ws, cell);
}

// 名前を付けて保存
// .xlsxを付けていないとエラーが出るTODO
// path パス
__declspec(dllexport) int saveAs(std::string path)
{
	XLUtil util;

	return util.saveAs(path);
}
