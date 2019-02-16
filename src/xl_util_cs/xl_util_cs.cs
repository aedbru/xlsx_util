using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// ClosedXML
using ClosedXML.Excel;

// 構造体用
using System.Runtime.InteropServices;


namespace xl_util_cs
{
    public struct Data
    {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string path;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string cell;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string val_c;
        //[MarshalAs(UnmanagedType.I4)]
        //public int val_i;
    }

    public class WrapClosedXML
    {
        static public XLWorkbook sBook;

        public int OpenWorkBook()
        {
            sBook = new XLWorkbook();

            return 0;
        }

        public int AddWorkSheet(int pData)
        {
            if (sBook == null) return -1;

            // 構造体に戻す
            System.IntPtr ptr = new System.IntPtr(pData);
            Data data = (Data)Marshal.PtrToStructure(ptr, typeof(Data));


            // パスは直接渡しても大丈夫かもしれない？TODO
            sBook.AddWorksheet(data.path);

            return 0;
        }

        public int SetValue(int pData)
        {
            if (sBook == null) return -1;

            // 構造体に戻す
            System.IntPtr ptr = new System.IntPtr(pData);
            Data data = (Data)Marshal.PtrToStructure(ptr, typeof(Data));

            sBook.Worksheet(data.path).Cell(data.cell).SetValue<string>(data.val_c);

            return 0;
        }

        public int GetValue(int pData)
        {
            if (sBook == null) return -1;

            // s32から構造体に戻す
            System.IntPtr ptr = new System.IntPtr(pData);
            Data data = (Data)Marshal.PtrToStructure(ptr, typeof(Data));


            // C++側に渡すための構造体の作成
            //Data retdata;
            //System.IntPtr retptr;
            data.val_c = sBook.Worksheet(data.path).Cell(data.cell).GetString();
            Marshal.StructureToPtr<Data>(data, ptr, true);

            return ptr.ToInt32();
        }

        public int SaveAs(int pData)
        {
            if (sBook == null) return -1;

            // 構造体に戻す
            System.IntPtr ptr = new System.IntPtr(pData);
            Data data = (Data)Marshal.PtrToStructure(ptr, typeof(Data));

            // 文字列からパスの作成


            sBook.SaveAs(data.path);

            return 0;
        }
    }
}
