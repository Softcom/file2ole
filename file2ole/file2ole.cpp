#include "stdafx.h"

using namespace std;

#define CheckHr(hr) if (FAILED(hr)) {return hr;}

HRESULT PackageOleObject(LPCTSTR inputFile, LPCTSTR outputFile)
{
	//Initialize COM
	HRESULT hr = CoInitialize(S_OK);
	CheckHr(hr);

	hr = OleInitialize(NULL);
	CheckHr(hr);

	IStoragePtr pStorage = NULL;
	IOleObjectPtr pOle = NULL;
	IDataObjectPtr pdo = NULL;
	FORMATETC fetc;
	STGMEDIUM stgm;
	HENHMETAFILE hmeta = NULL;

	// Create a compound storage document.
	hr = StgCreateStorageEx(
		outputFile,
		STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_TRANSACTED,
		STGFMT_DOCFILE,
		0,
		NULL,
		NULL,
		IID_IStorage,
		reinterpret_cast<void**>(&pStorage));
	CheckHr(hr);

	// Create OLE package from file.
	hr = OleCreateFromFile(CLSID_NULL, inputFile, ::IID_IOleObject, OLERENDER_NONE, NULL, NULL, pStorage, (void**)&pOle);
	CheckHr(hr);

	hr = OleRun(pOle);
	CheckHr(hr);

	hr = pOle->QueryInterface(IID_IDataObject, (void**)&pdo);
	CheckHr(hr);

	fetc.cfFormat = CF_ENHMETAFILE;
	fetc.dwAspect = DVASPECT_CONTENT;
	fetc.lindex = -1;
	fetc.ptd = NULL;
	fetc.tymed = TYMED_ENHMF;

	stgm.hEnhMetaFile = NULL;
	stgm.tymed = TYMED_ENHMF;
	hr = pdo->GetData(&fetc, &stgm);
	CheckHr(hr);

	// Create image metafile for object.
	//CopyEnhMetaFile(stgm.hEnhMetaFile, emfFile);

	hr = pStorage->Commit(STGC_DEFAULT);
	CheckHr(hr);

	pOle->Close(0);
	DeleteEnhMetaFile(stgm.hEnhMetaFile);
	DeleteEnhMetaFile(hmeta);

	//-->leads to exception
	//OleUninitialize();

	CoUninitialize();

	return hr;
}

wchar_t *convertCharArrayToLPCWSTR(const char* charArray)
{
	wchar_t* wString = new wchar_t[4096];
	MultiByteToWideChar(CP_ACP, 0, charArray, -1, wString, 4096);
	return wString;
}

int main(int argc, char* argv[]) {

	if (argc != 3) {
		printf("Invalid number of parameters!\nUsage: file2ole <inputFile> <outputFile>");
		return 1;
	}

	FILE* file = NULL;
	int error = error = fopen_s(&file, argv[1], "r");
	if (!error) {
		fclose(file);
	}
	else {
		printf("Input file can not be read!");
		return error;
	}

	LPCTSTR input = convertCharArrayToLPCWSTR(argv[1]);
	LPCTSTR output = convertCharArrayToLPCWSTR(argv[2]);

	printf("Parameters:\nIn: %S\nOut: %S", input, output);
	
	//-->for debugging
	//getchar();

	error = PackageOleObject(input, output);
	if (error) {
		printf("Error while processing: HRESULT (0x%8X)", error);
	}
	return error;
}