//#define CURL_STATICLIB

#include <windows.h>
#include <commctrl.h>
#include <process.h>

#include <iostream>
#include <sstream>
#include <fstream>
#include <vector>

#include "zlibengn.h"
#include "curl/curl.h"

#include <stdlib.h>

ZlibEngine ZipEngine;

using namespace std;


void __cdecl DownloadURL(void *threadarg);
void __cdecl UpdateThread(void *threadarg);

std::string FileURL;
std::string TheFile;
boolean DownloadComplete;

void *pointer;

size_t my_write_func(void *ptr, size_t size, size_t nmemb, FILE *stream);
size_t my_read_func(void *ptr, size_t size, size_t nmemb, FILE *stream);
void LogAnError(string error);

const int ID_UPDATER = 1;
const int ID_PROGRESS = 2;
const int IDC_PROGRESS1 = 1	;

//---------------------------------------------------------------------------
HWND hWnd;
HINSTANCE hInst;
LRESULT CALLBACK DlgProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);
//---------------------------------------------------------------------------
INT WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
				   LPSTR lpCmdLine, int nCmdShow)
{
	FileURL.clear();
	TheFile.clear();

	try
	{
		hInst = hInstance;

		DialogBox(hInst, MAKEINTRESOURCE(ID_UPDATER),
				hWnd, reinterpret_cast<DLGPROC>(DlgProc));
	}
	catch (...)
	{
		LogAnError("Error updating");
	}

	ShellExecute(NULL, "open", "ody.exe", NULL, NULL, SW_SHOWNORMAL);

	return FALSE;
}
//---------------------------------------------------------------------------
LRESULT CALLBACK DlgProc(HWND hWndDlg, UINT Msg,
		       WPARAM wParam, LPARAM lParam)
{
	INITCOMMONCONTROLSEX InitCtrlEx;

	InitCtrlEx.dwSize = sizeof(INITCOMMONCONTROLSEX);
	InitCtrlEx.dwICC  = ICC_PROGRESS_CLASS;
	InitCommonControlsEx(&InitCtrlEx);

	switch(Msg)
	{
	case WM_INITDIALOG:
		hWnd = hWndDlg;
		_beginthread(UpdateThread, 0, NULL);
		return TRUE;

	case WM_COMMAND:
		switch(wParam)
		{
		case IDOK:
			EndDialog(hWndDlg, 0);
			return TRUE;
		case IDCANCEL:
			EndDialog(hWndDlg, 0);
			return TRUE;
		}
		break;
	}

	return FALSE;
}

size_t my_write_func(void *ptr, size_t size, size_t nmemb, FILE *stream)
{
  return fwrite(ptr, size, nmemb, stream);
}

size_t my_read_func(void *ptr, size_t size, size_t nmemb, FILE *stream)
{
  return fread(ptr, size, nmemb, stream);
}
 
void setText(std::string Text)
{
	SetDlgItemText(hWnd, ID_PROGRESS, Text.c_str());
}


int my_progress_func(void *notused, double t, double d, double ultotal, double ulnow)
{
	//Bar->SetValue((int)(d*100.0/t));

	SendDlgItemMessage(hWnd, IDC_PROGRESS1, PBM_SETRANGE, 0L, MAKELPARAM(0, 100));
	SendDlgItemMessage(hWnd, IDC_PROGRESS1, PBM_SETPOS, (long)(d*100.0/t), 0);

	string OutputString;
	OutputString = "Updating ";
	OutputString += TheFile;
	OutputString += " - ";
	std::ostringstream convert;
	convert << (int)(d*100.0/t);
	OutputString += convert.str();
	OutputString += "%";

	setText(OutputString);

	ulnow = 0;
	ultotal = 0;
	return 0;
}
void __cdecl DownloadURL(void *threadarg)
{
  CURL *curl;
  CURLcode res;
  FILE *outfile;

  remove(TheFile.c_str());

  curl = curl_easy_init();
  if(curl)
  {
		outfile = fopen(TheFile.c_str(), "wb");

		if (outfile != NULL)
		{
			curl_easy_setopt(curl, CURLOPT_URL, FileURL.c_str());
			curl_easy_setopt(curl, CURLOPT_WRITEDATA, outfile);
			curl_easy_setopt(curl, CURLOPT_WRITEFUNCTION, my_write_func);
			curl_easy_setopt(curl, CURLOPT_READFUNCTION, my_read_func);
			curl_easy_setopt(curl, CURLOPT_NOPROGRESS, FALSE);
			curl_easy_setopt(curl, CURLOPT_PROGRESSFUNCTION, my_progress_func);
			curl_easy_setopt(curl, CURLOPT_PROGRESSDATA, NULL);
			curl_easy_setopt(curl, CURLOPT_FAILONERROR, 1);

			res = curl_easy_perform(curl);

			fclose(outfile);
		}
		else
		{
			std::string ErrorMessage;
			ErrorMessage = "Cannot open the file ";
			ErrorMessage +=  TheFile.c_str();
			ErrorMessage += " for writing.  The file is in use.  Please close the game first to prevent this problem.";
			MessageBox(NULL, ErrorMessage.c_str(), "Updating Error", 0);
		}

		curl_easy_cleanup(curl);

		threadarg = 0;
  }

  DownloadComplete = true;
  return;
}

void __cdecl UpdateThread(void *threadarg)
{
	remove ("update.dat");

	string SiteURL = "http://www.codemallet.com/odyssey/updates/";
	TheFile = "update.dat";
	FileURL = SiteURL;
	FileURL += TheFile;
	DownloadComplete = false;
	_beginthread(DownloadURL, 0, NULL);
	while (DownloadComplete == false) Sleep(1);

	ifstream updatedat ("update.dat");

	if (updatedat.is_open())
	{
		while (!updatedat.eof())
		{
			std::getline(updatedat, TheFile);
			if (TheFile.length() > 0)
			{
				// Split the file - Delimiter is a comma
				vector < std::string > parseline;
				std::string  temp;

				while (TheFile.find(",", 0) != std::string::npos)
				{
					size_t  pos = TheFile.find(",", 0);
					temp = TheFile.substr(0, pos);
					TheFile.erase(0, pos + 1);
					parseline.push_back(temp);
				}

				parseline.push_back(TheFile);

				TheFile = parseline.front();

				bool DownloadUpdate = true;

				if (TheFile.length() > 0)
				{
					// Check if the file currently exists and it is the right size
					FILE *fin;
					FILE *fout;
					long length;
					fin = fopen( TheFile.c_str(), "rb" );

					if (fin != NULL)
					{
						length = _filelength( _fileno( fin ) );

						unsigned char *buf = new unsigned char[length];

						fread(buf, 1, length, fin);

						long CRC2 = 0;
						CRC2 = crc32(CRC2, buf, length);

						string tempstring = parseline.at(2);
						long length2 = atol(tempstring.c_str());

						tempstring = parseline.at(3);
						long CRC = atol(tempstring.c_str());


						if (length == length2)
						{
							if (CRC == CRC2)
							{
								DownloadUpdate = false;
							}
							else
								DownloadUpdate = true;
						}
						else
						{
							DownloadUpdate = true;
						}

						fclose( fin );
					}
					else
						DownloadUpdate = true;

					fin = 0;
					fout = 0;
				}
				else
					DownloadUpdate = true;

				if (DownloadUpdate == true)
				{
					// Download the latest version

					DownloadComplete = false;
					setText(std::string("Updating " + TheFile + " - 0%"));

					FileURL = SiteURL;
					FileURL += TheFile;

					_beginthread(DownloadURL, 0, NULL);
					while (DownloadComplete == false) Sleep(1);

					// The following is a nasty hack to compare the file sizes using stringstreams
					ostringstream tempstringstream;
					ostringstream tempstringstream2;

					string tempstring = parseline.at(1);
					tempstringstream << tempstring;

					ifstream tempfile (TheFile.c_str());
					tempfile.seekg(0,ios_base::end);
					int size = tempfile.tellg();
					tempstringstream2 << size;
					tempfile.close();

					if (tempstringstream.str() == tempstringstream2.str())
					{
						
					}
					else
					{
						string Error = "Invalid Pre-Decompression Size:  Error updating ";
						Error += TheFile;
						LogAnError(Error);
					}

					// Decompress the File
					remove("temp");
					ZipEngine.decompress(TheFile.c_str(), "temp");
					remove(TheFile.c_str());
					rename("temp", TheFile.c_str());

					// The following is a nasty hack to compare the file sizes using stringstreams
					tempstringstream.clear();
					tempstringstream2.clear();

					tempstring = parseline.at(2);
					tempstringstream << tempstring;

					tempfile.open (TheFile.c_str());
					tempfile.seekg(0,ios_base::end);
					size = tempfile.tellg();
					tempstringstream2 << size;
					tempfile.close();

					if (tempstringstream.str() == tempstringstream2.str())
					{
						
					}
					else
					{
						string Error = "Invalid Post-Decompression Size:  Error updating ";
						Error += TheFile;
						LogAnError(Error);
					}
				}
			}
		}

		updatedat.close();
	}
	else
	{
		LogAnError("Unable to open update.dat");
	}

	EndDialog(hWnd, 0);

	threadarg = 0;
	return;
}

void CreateMainWindow(HINSTANCE hInstance)
{
    WNDCLASSEX wc;
    HWND hwnd;
    MSG Msg;

    //Step 1: Registering the Window Class
    wc.cbSize        = sizeof(WNDCLASSEX);
    wc.style         = 0;
    wc.lpfnWndProc   = DlgProc;
    wc.cbClsExtra    = 0;
    wc.cbWndExtra    = 0;
    wc.hInstance     = hInstance;
    wc.hIcon         = LoadIcon(NULL, IDI_APPLICATION);
    wc.hCursor       = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
    wc.lpszMenuName  = NULL;
    wc.lpszClassName = "CodeMallet Updater";
    wc.hIconSm       = LoadIcon(NULL, IDI_APPLICATION);

    if(!RegisterClassEx(&wc))
    {
        MessageBox(NULL, "Window Registration Failed!", "Error!",
            MB_ICONEXCLAMATION | MB_OK);
        return;
    }

    // Step 2: Creating the Window
    hwnd = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        "CodeMallet Updater",
        "CodeMallet Updater",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 240, 120,
        NULL, NULL, hInstance, NULL);

    if(hwnd == NULL)
    {
        MessageBox(NULL, "Window Creation Failed!", "Error!",
            MB_ICONEXCLAMATION | MB_OK);
        return;
    }

    ShowWindow(hwnd, SW_SHOWNORMAL);
    UpdateWindow(hwnd);

    // Step 3: The Message Loop
    while(GetMessage(&Msg, NULL, 0, 0) > 0)
    {
        TranslateMessage(&Msg);
        DispatchMessage(&Msg);
    }
    return;
}

void __cdecl LogAnError(string error)
{
	time_t mytime;
	struct tm *today;
	char ftime[30];
	char fdate[30];
	
	time(&mytime);
	today = localtime(&mytime);
	strftime(ftime, 30, "%Y-%b-%d.%H.%M.%S", today );
	strftime(fdate, 30, "%Y-%b-%d", today );

	string filepath;
	filepath = fdate;
	filepath += ".log";

	FILE *file;
	file = fopen(filepath.c_str(),"a");
	char linebreak[3];
	linebreak[0] = 13;
	linebreak[1] = 10;
	linebreak[2] = 0;

	string output;
	output += linebreak;
	output += ftime;
	output += " - " + error;

	fwrite(output.c_str(), output.size(), 1, file); 

	fclose(file);
}
