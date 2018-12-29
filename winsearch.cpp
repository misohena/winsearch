/**
 * @file
 * @brief Windows Search command-line client Visual C++ port
 * @author AKIYAMA Kouhei
 * @since 2010-07-01
 *
 * Original python script is:
 * http://www.carcosa.net/jason/software/utilities/desktopsearch/  (Jason F. McBrayer)
 * 
 */
#import "c:\Program Files\Common Files\System\ADO\msado15.dll"  rename("EOF", "ADO_EOF")
#include <initguid.h>
#include <adoid.h>
#include <adoint.h>

#include <atlbase.h>
#include <windows.h>
#include <tchar.h>

#include <string>
#include <clocale>
#include <cstdio>

//typedef ADODB::_ConnectionPtr ADOConnectionPtr;
//typedef ADODB::_RecordsetPtr ADORecordsetPtr;
//typedef ADODB::ErrorPtr ADOErrorPtr;
//typedef ADODB::_CommandPtr ADOCommandPtr;
//typedef ADODB::_ParameterPtr ADOParameterPtr;
//typedef ADODB::FieldPtr ADOFieldPtr;

// main
int _tmain(int argc, TCHAR *argv[])
{
    const int maxResultCount = 30;
    if(argc <= 1){
        _tprintf(_T("Usage: winsearch keyword...\n"));
        return -1;
    }

    // set locale
    setlocale(LC_ALL, "");

    // create query
    static const TCHAR * const COLUMN_FILE_DIR = _T("System.ItemFolderPathDisplay");
    static const TCHAR * const COLUMN_FILE_NAME = _T("System.FileName");
    using string = std::basic_string<TCHAR>;

    string query_names;
    string query_contains;
    for(int ai = 1; ai < argc; ++ai){
        if(ai > 1){
            query_names += _T(" AND ");
            query_contains += _T(" AND ");
        }
        query_names += _T("System.FileName Like '%") + string(argv[ai]) + _T("%'"); ///@todo escape
        query_contains += _T("Contains('\"") + string(argv[ai]) + _T("\"')");  ///@todo escape
    }
    string query
        = string(_T("SELECT "))
            + COLUMN_FILE_DIR + _T(",")
            + COLUMN_FILE_NAME
        + _T(" FROM SYSTEMINDEX")
        + _T(" WHERE (")
            + query_names + _T(" ) OR ( ") + query_contains + _T(" )");

    // initialize COM
    ::CoInitialize(NULL);
    // search
    ADODB::_ConnectionPtr conn;
    try{
        if(conn.CreateInstance(__uuidof(ADODB::Connection)) == S_OK){
            // open
            conn->Open(
                _T("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"),
                _T(""), //UserID
                _T(""), //Password
                NULL);

            // execute
            ADODB::_RecordsetPtr recordset = conn->Execute(query.c_str(), NULL, ADODB::adCmdText);

            // enumerate
            if(! (recordset->GetBOF() && recordset->GetADO_EOF()) ){
                recordset->MoveFirst();

                int count = 0;
				const _variant_t varColumnFileDir = COLUMN_FILE_DIR;
				const _variant_t varColumnFileName = COLUMN_FILE_NAME;
                while( ! recordset->ADO_EOF && count < maxResultCount){
                    _tprintf(_T("%s\\%s\n"),
                        (recordset->Fields->GetItem(varColumnFileDir )->GetValue()).bstrVal,
                        (recordset->Fields->GetItem(varColumnFileName)->GetValue()).bstrVal
                    );
                    recordset->MoveNext();
                    ++count;
                }
            }

        }
        else{
			_ftprintf(stderr, _T("Error: Failed to connect database\n"));
        }
    }catch(_com_error &e){
        _ftprintf(stderr, _T("Error: (0x%08x):%s\n"), e.Error() , (LPWSTR)e.Description());
    }

    // close
    if(conn && conn->State != ADODB::adStateClosed){
        conn->Close();
    }

    ::CoUninitialize();

    return 0;
}
