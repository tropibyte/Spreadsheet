#pragma once

class CGrid32Mgr;

class CFormulator
{
	friend class CGrid32Mgr;

public:
	// Needed by file-scope text/date function helpers; safe to expose.
	static bool ParseCellRef(const std::wstring& token, UINT& row, UINT& col);
private:
	static void SkipSpaces(const std::wstring& s, size_t& pos);
	static bool IsValidCellReference(const std::wstring& s, size_t& pos, size_t& row, size_t& col);
	static double GetCellValue(CGrid32Mgr* mgr, UINT row, UINT col,
		std::set<std::pair<UINT, UINT>>& visited);

	static double ParseFactor(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos,
		std::set<std::pair<UINT, UINT>>& visited);

	static double SumRange(CGrid32Mgr* mgr, const std::wstring& arg,
		std::set<std::pair<UINT, UINT>>& visited);
	static double MinMaxRange(CGrid32Mgr* mgr, const std::wstring& arg,
		std::set<std::pair<UINT, UINT>>& visited, bool bMax);
	static double CountRange(CGrid32Mgr* mgr, const std::wstring& arg,
		std::set<std::pair<UINT, UINT>>& visited);
	static double ParseTerm(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, std::set<std::pair<UINT, UINT>>& visited);
	static double ParseExpression(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, std::set<std::pair<UINT, UINT>>& visited);
public:
	// Used by both internal helpers and (transitively) by file-scope text
	// function helpers — promoted so the namespace can call back in.
	static double EvalArg(CGrid32Mgr* mgr, const std::wstring& s, std::set<std::pair<UINT, UINT>>& visited);
};
