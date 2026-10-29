#pragma once

class CGrid32Mgr;

class CFormulator
{
	friend class CGrid32Mgr;

public:
	// Needed by file-scope text/date function helpers; safe to expose.
	// Accepts a *whole* reference token only ("A1"); anything with trailing
	// characters ("A1:B10", "A1B") is rejected so range syntax can't be
	// mistaken for a single cell.
	static bool ParseCellRef(const std::wstring& token, UINT& row, UINT& col);
	// Parse "A1" or "A1:B10" into an inclusive, normalized rectangular range.
	static bool ParseRange(const std::wstring& arg, UINT& sr, UINT& sc,
		UINT& er, UINT& ec);
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
	static double AverageRange(CGrid32Mgr* mgr, const std::wstring& arg,
		std::set<std::pair<UINT, UINT>>& visited);
	static double ParseTerm(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, std::set<std::pair<UINT, UINT>>& visited);
	// Additive level: term (('+'|'-') term)*. ParseExpression sits above this
	// and adds the comparison operators, which bind least tightly.
	static double ParseAdditive(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, std::set<std::pair<UINT, UINT>>& visited);
	static double ParseExpression(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, std::set<std::pair<UINT, UINT>>& visited);
public:
	// Used by both internal helpers and (transitively) by file-scope text
	// function helpers — promoted so the namespace can call back in.
	static double EvalArg(CGrid32Mgr* mgr, const std::wstring& s, std::set<std::pair<UINT, UINT>>& visited);
};
