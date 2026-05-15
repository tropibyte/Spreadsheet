// This MFC Samples source code demonstrates using MFC Microsoft Office Fluent User Interface
// (the "Fluent UI") and is provided only as referential material to supplement the
// Microsoft Foundation Classes Reference and related electronic documentation
// included with the MFC C++ library software.
// License terms to copy, use or distribute the Fluent UI are available separately.
// To learn more about our Fluent UI licensing program, please visit
// https://go.microsoft.com/fwlink/?LinkId=238214.
//
// Copyright (C) Microsoft Corporation
// All rights reserved.

// SpreadsheetDoc.cpp : implementation of the CSpreadsheetDoc class
//

#include "pch.h"
#include "framework.h"
// SHARED_HANDLERS can be defined in an ATL project implementing preview, thumbnail
// and search filter handlers and allows sharing of document code with that project.
#ifndef SHARED_HANDLERS
#include "Spreadsheet.h"
#endif

#include "SpreadsheetDoc.h"
#include "GridView.h"

#include <propkey.h>
#include <vector>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CSpreadsheetDoc

IMPLEMENT_DYNCREATE(CSpreadsheetDoc, CDocument)

BEGIN_MESSAGE_MAP(CSpreadsheetDoc, CDocument)
END_MESSAGE_MAP()


// CSpreadsheetDoc construction/destruction

CSpreadsheetDoc::CSpreadsheetDoc() noexcept
{
	// TODO: add one-time construction code here

}

CSpreadsheetDoc::~CSpreadsheetDoc()
{
}

BOOL CSpreadsheetDoc::OnNewDocument()
{
	if (!CDocument::OnNewDocument())
		return FALSE;

	// TODO: add reinitialization code here
	// (SDI documents will reuse this document)

	return TRUE;
}




// CSpreadsheetDoc serialization

void CSpreadsheetDoc::Serialize(CArchive& ar)
{
	// SSF v2 is line-based UTF-16 text; we bypass CArchive's binary
	// serialization via OnOpenDocument / OnSaveDocument instead. Serialize
	// is left as a stub so the framework's default flow doesn't error.
	UNREFERENCED_PARAMETER(ar);
}

// Locate the grid control owned by this document's view. Returns nullptr if
// the view hasn't been instantiated yet (shouldn't happen during normal
// open/save flow, but guards against framework edge cases).
static CGridCtrl* GetGridCtrl(CSpreadsheetDoc* pDoc)
{
	POSITION pos = pDoc->GetFirstViewPosition();
	if (!pos) return nullptr;
	CGridView* pView = DYNAMIC_DOWNCAST(CGridView, pDoc->GetNextView(pos));
	if (!pView) return nullptr;
	return &pView->GetGridCtrl();
}

BOOL CSpreadsheetDoc::OnOpenDocument(LPCTSTR lpszPathName)
{
	if (!CDocument::OnOpenDocument(lpszPathName))
		return FALSE;

	CFile file;
	if (!file.Open(lpszPathName, CFile::modeRead | CFile::typeBinary))
		return FALSE;

	ULONGLONG sz = file.GetLength();
	std::vector<wchar_t> buffer;

	// Detect UTF-16 LE BOM. Files without BOM are read as raw wchar_t.
	if (sz >= 2)
	{
		unsigned char bom[2] = { 0 };
		file.Read(bom, 2);
		if (bom[0] == 0xFF && bom[1] == 0xFE)
			sz -= 2;
		else
			file.Seek(0, CFile::begin);
	}

	size_t nWChars = (size_t)(sz / sizeof(wchar_t));
	buffer.resize(nWChars + 1, L'\0');
	if (nWChars > 0)
		file.Read(buffer.data(), (UINT)(nWChars * sizeof(wchar_t)));
	file.Close();

	CGridCtrl* pGrid = GetGridCtrl(this);
	if (!pGrid)
		return FALSE;

	GCSTREAM stream{};
	stream.m_cbSize = sizeof(GCSTREAM);
	stream.m_pwszBuff = buffer.data();
	stream.m_cbBuffSize = (UINT)((nWChars + 1) * sizeof(wchar_t));
	stream.m_dwFormat = SF_SSF;
	pGrid->StreamIn(stream);
	return stream.m_dwError == 0;
}

BOOL CSpreadsheetDoc::OnSaveDocument(LPCTSTR lpszPathName)
{
	CGridCtrl* pGrid = GetGridCtrl(this);
	if (!pGrid)
		return FALSE;

	// StreamOut writes into a caller-provided buffer; if it doesn't fit, it
	// returns GRID_ERROR_OUT_OF_RANGE. Grow geometrically up to a sane cap.
	std::vector<wchar_t> buffer(64 * 1024, L'\0');
	GCSTREAM stream{};
	for (int attempt = 0; attempt < 8; ++attempt)
	{
		stream = GCSTREAM{};
		stream.m_cbSize = sizeof(GCSTREAM);
		stream.m_pwszBuff = buffer.data();
		stream.m_cbBuffSize = (UINT)(buffer.size() * sizeof(wchar_t));
		stream.m_dwFormat = SF_SSF;
		pGrid->StreamOut(stream);
		if (stream.m_dwError == 0)
			break;
		if (stream.m_dwError != GRID_ERROR_OUT_OF_RANGE)
			return FALSE;
		buffer.resize(buffer.size() * 2);
	}
	if (stream.m_dwError != 0)
		return FALSE;

	size_t actualWChars = wcsnlen(buffer.data(), buffer.size());

	CFile file;
	if (!file.Open(lpszPathName, CFile::modeCreate | CFile::modeWrite | CFile::typeBinary))
		return FALSE;
	const unsigned char bom[2] = { 0xFF, 0xFE };
	file.Write(bom, 2);
	if (actualWChars > 0)
		file.Write(buffer.data(), (UINT)(actualWChars * sizeof(wchar_t)));
	file.Close();
	return TRUE;
}

#ifdef SHARED_HANDLERS

// Support for thumbnails
void CSpreadsheetDoc::OnDrawThumbnail(CDC& dc, LPRECT lprcBounds)
{
	// Modify this code to draw the document's data
	dc.FillSolidRect(lprcBounds, RGB(255, 255, 255));

	CString strText = _T("TODO: implement thumbnail drawing here");
	LOGFONT lf;

	CFont* pDefaultGUIFont = CFont::FromHandle((HFONT) GetStockObject(DEFAULT_GUI_FONT));
	pDefaultGUIFont->GetLogFont(&lf);
	lf.lfHeight = 36;

	CFont fontDraw;
	fontDraw.CreateFontIndirect(&lf);

	CFont* pOldFont = dc.SelectObject(&fontDraw);
	dc.DrawText(strText, lprcBounds, DT_CENTER | DT_WORDBREAK);
	dc.SelectObject(pOldFont);
}

// Support for Search Handlers
void CSpreadsheetDoc::InitializeSearchContent()
{
	CString strSearchContent;
	// Set search contents from document's data.
	// The content parts should be separated by ";"

	// For example:  strSearchContent = _T("point;rectangle;circle;ole object;");
	SetSearchContent(strSearchContent);
}

void CSpreadsheetDoc::SetSearchContent(const CString& value)
{
	if (value.IsEmpty())
	{
		RemoveChunk(PKEY_Search_Contents.fmtid, PKEY_Search_Contents.pid);
	}
	else
	{
		CMFCFilterChunkValueImpl *pChunk = nullptr;
		ATLTRY(pChunk = new CMFCFilterChunkValueImpl);
		if (pChunk != nullptr)
		{
			pChunk->SetTextValue(PKEY_Search_Contents, value, CHUNK_TEXT);
			SetChunkValue(pChunk);
		}
	}
}

#endif // SHARED_HANDLERS

// CSpreadsheetDoc diagnostics

#ifdef _DEBUG
void CSpreadsheetDoc::AssertValid() const
{
	CDocument::AssertValid();
}

void CSpreadsheetDoc::Dump(CDumpContext& dc) const
{
	CDocument::Dump(dc);
}
#endif //_DEBUG


// CSpreadsheetDoc commands
