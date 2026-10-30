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
	// SDI reuses this document for File>New. CDocument::OnNewDocument calls
	// DeleteContents (below), which empties the grid and its undo/redo
	// history, and then marks the document clean.
	if (!CDocument::OnNewDocument())
		return FALSE;

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

// MFC calls DeleteContents from CDocument::OnNewDocument and
// CDocument::OnOpenDocument (and when the document closes), so both File>New
// and File>Open start from an empty grid with empty undo and redo stacks.
// Without it the SDI document kept the old sheet on New, and Open merged the
// file into whatever the grid already held (AUDIT_2026-09.md H15).
void CSpreadsheetDoc::DeleteContents()
{
	// On close the views are destroyed before this runs; nothing to clear then.
	CGridCtrl* pGrid = GetGridCtrl(this);
	if (pGrid && ::IsWindow(pGrid->GetSafeHwnd()))
		pGrid->ClearAll();
	CDocument::DeleteContents();
}

// Render the grid as SSF into `buffer`, returning the grid's error code and,
// on success, the number of bytes written in `cbOut`. At most two passes: the
// first into a buffer that fits ordinary sheets, the second -- only if that
// was too small -- into one sized exactly from what the grid says it needs.
// The old loop doubled the buffer and re-rendered the whole sheet up to eight
// times, and still failed any sheet larger than 16 MB (AUDIT_2026-09.md P1).
static DWORD StreamOutSsf(CGridCtrl& grid, std::vector<wchar_t>& buffer, UINT& cbOut)
{
	buffer.assign(64 * 1024, L'\0');
	for (int attempt = 0; attempt < 2; ++attempt)
	{
		GCSTREAM stream{};
		stream.m_cbSize = sizeof(GCSTREAM);
		stream.m_pwszBuff = buffer.data();
		stream.m_cbBuffSize = (UINT)(buffer.size() * sizeof(wchar_t));
		stream.m_dwFormat = SF_SSF;
		grid.StreamOut(stream);
		cbOut = stream.m_cbBuffOut;
		if (stream.m_dwError != GRID_ERROR_OUT_OF_RANGE)
			return stream.m_dwError;
		// m_cbBuffOut is the size needed, terminator included; 0 means the
		// output cannot be described in a UINT at all.
		size_t needed = stream.m_cbBuffOut / sizeof(wchar_t);
		if (needed <= buffer.size())
			break;
		buffer.assign(needed, L'\0');
	}
	return GRID_ERROR_OUT_OF_RANGE;
}

BOOL CSpreadsheetDoc::OnOpenDocument(LPCTSTR lpszPathName)
{
	// The base opens the file to check it can be read, then calls
	// DeleteContents, which empties the grid. Every failure after this point
	// therefore leaves the document modified: CSingleDocTemplate then resets
	// to a fresh untitled document instead of keeping the old file's name on
	// an emptied grid, where the next Save would overwrite that file.
	if (!CDocument::OnOpenDocument(lpszPathName))
		return FALSE;

	std::vector<wchar_t> buffer;
	size_t nWChars = 0;
	try
	{
		CFile file(lpszPathName, CFile::modeRead | CFile::shareDenyWrite | CFile::typeBinary);
		ULONGLONG sz = file.GetLength();

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

		nWChars = (size_t)(sz / sizeof(wchar_t));
		buffer.resize(nWChars + 1, L'\0');
		if (nWChars > 0)
			file.Read(buffer.data(), (UINT)(nWChars * sizeof(wchar_t)));
		file.Close();
	}
	catch (CException* e)
	{
		ReportSaveLoadException(lpszPathName, e, FALSE, AFX_IDP_FAILED_TO_OPEN_DOC);
		e->Delete();
		SetModifiedFlag(TRUE);
		return FALSE;
	}

	CGridCtrl* pGrid = GetGridCtrl(this);
	if (!pGrid)
	{
		SetModifiedFlag(TRUE);
		return FALSE;
	}

	// StreamIn replaces the grid's content and leaves undo/redo empty; if the
	// file does not parse, the grid is left as DeleteContents left it.
	GCSTREAM stream{};
	stream.m_cbSize = sizeof(GCSTREAM);
	stream.m_pwszBuff = buffer.data();
	stream.m_cbBuffSize = (UINT)((nWChars + 1) * sizeof(wchar_t));
	stream.m_dwFormat = SF_SSF;
	pGrid->StreamIn(stream);
	if (stream.m_dwError != 0)
	{
		ReportSaveLoadException(lpszPathName, nullptr, FALSE, AFX_IDP_FAILED_TO_OPEN_DOC);
		SetModifiedFlag(TRUE);
		return FALSE;
	}

	// Loading fires GN_CONTENTCHANGED, which marks the document dirty; a
	// freshly opened file is not modified.
	SetModifiedFlag(FALSE);
	return TRUE;
}

BOOL CSpreadsheetDoc::OnSaveDocument(LPCTSTR lpszPathName)
{
	CGridCtrl* pGrid = GetGridCtrl(this);
	if (!pGrid)
		return FALSE;

	// Render first, so a failure here never truncates the file on disk.
	std::vector<wchar_t> buffer;
	UINT cbData = 0;
	if (StreamOutSsf(*pGrid, buffer, cbData) != 0)
	{
		ReportSaveLoadException(lpszPathName, nullptr, TRUE, AFX_IDP_FAILED_TO_SAVE_DOC);
		return FALSE;
	}

	try
	{
		CFile file(lpszPathName, CFile::modeCreate | CFile::modeWrite | CFile::shareExclusive | CFile::typeBinary);
		const unsigned char bom[2] = { 0xFF, 0xFE };
		file.Write(bom, 2);
		// m_cbBuffOut: exactly the bytes the grid wrote, terminator excluded.
		if (cbData > 0)
			file.Write(buffer.data(), cbData);
		file.Close();
	}
	catch (CException* e)
	{
		ReportSaveLoadException(lpszPathName, e, TRUE, AFX_IDP_FAILED_TO_SAVE_DOC);
		e->Delete();
		return FALSE;
	}

	// CDocument::OnSaveDocument would normally clear this, but it is not
	// called: the base writes the file itself through Serialize(), which is a
	// stub here, so invoking it would truncate what we just wrote.
	SetModifiedFlag(FALSE);
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

// Text still being typed into a cell is part of the document. It used to be
// committed only when the editor lost focus, which happens after
// SaveModified has already decided the document is clean, so closing or
// saving mid-edit dropped it (AUDIT_2026-10 S5). Committing raises
// GN_CONTENTCHANGED, which sets the modified flag before the checks below.
void CSpreadsheetDoc::CommitPendingEdits()
{
	POSITION pos = GetFirstViewPosition();
	while (pos != NULL)
	{
		if (CGridView* pView = DYNAMIC_DOWNCAST(CGridView, GetNextView(pos)))
			pView->CommitEdit();
	}
}

BOOL CSpreadsheetDoc::SaveModified()
{
	CommitPendingEdits();
	return CDocument::SaveModified();
}

BOOL CSpreadsheetDoc::DoSave(LPCTSTR lpszPathName, BOOL bReplace)
{
	CommitPendingEdits();
	return CDocument::DoSave(lpszPathName, bReplace);
}
